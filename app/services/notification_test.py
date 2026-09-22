"""通知规则的"测试发送"服务。

`app/reply_server.py` 的 ``POST /message-notifications/rule/{rule_id}/test``
依赖本模块：管理员点"测试发送"时，必须**真实**往规则绑定的渠道发一条消息，
并把渠道返回的成功/失败如实翻译成 HTTP 语义。

与运行时的通知发送（`XianyuLive` 内部会吞掉第三方错误）不同，这里的契约是
一次调用给出明确结论，因此：

* 规则或渠道不存在 → 404；
* 发送过于频繁 → 429，并带上 ``Retry-After`` 秒数；
* 渠道配置错误 → 400（错误信息里不回显密钥、地址里的查询参数）；
* 第三方拒绝/超时/网络异常 → 502，前端提示可重试。

本文件在上游公开源码里缺失（作者漏提交），这里按调用点
（`reply_server.py` 第 33-37 行导入、第 3446-3480 行消费）与既有实现
（`app.services.notification_channels`、`app.services.notification_sender`）
对齐补齐。
"""

from __future__ import annotations

import asyncio
import time
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Optional

from loguru import logger

from .notification_channels import (
    NotificationChannelConfigError,
    normalize_channel_type,
)
from .notification_sender import (
    NotificationSendError,
    NotificationSendTimeout,
    NotificationSender,
)

__all__ = [
    "NotificationTestError",
    "NotificationTestRateLimiter",
    "notification_test_rate_limiter",
    "NotificationTestService",
]

DEFAULT_RATE_LIMIT_SECONDS = 10.0
MAX_RULE_NAME_CHARS = 60


class NotificationTestError(Exception):
    """测试发送失败：同时携带 HTTP 状态码与可选的 ``Retry-After`` 秒数。"""

    def __init__(
        self,
        code: str,
        message: str,
        *,
        status_code: int = 400,
        retry_after: Optional[int] = None,
        category: str = "config",
    ) -> None:
        super().__init__(message)
        self.code = str(code)
        self.message = str(message)
        self.status_code = int(status_code)
        self.retry_after = int(retry_after) if retry_after else None
        self.category = str(category)

    def detail(self) -> Dict[str, Any]:
        """前端错误解析器读取的形状：``detail.message``。"""
        payload: Dict[str, Any] = {
            "code": self.code,
            "message": self.message,
            "category": self.category,
        }
        if self.retry_after:
            payload["retry_after"] = self.retry_after
        return payload


class NotificationTestRateLimiter:
    """按用户限制测试发送频率，避免拿真实渠道做压力测试。"""

    def __init__(
        self,
        interval_seconds: float = DEFAULT_RATE_LIMIT_SECONDS,
        *,
        max_users: int = 1024,
        clock: Any = None,
    ) -> None:
        self.interval_seconds = max(0.0, float(interval_seconds))
        self.max_users = max(1, int(max_users))
        self._clock = clock or time.monotonic
        self._last_sent: Dict[str, float] = {}
        self._lock = asyncio.Lock()

    async def acquire(self, user_id: Any) -> None:
        """允许发送则记录时间戳；否则抛出带 ``Retry-After`` 的错误。"""
        if self.interval_seconds <= 0:
            return

        key = str(user_id)
        async with self._lock:
            now = self._clock()
            last = self._last_sent.get(key)
            if last is not None:
                remaining = self.interval_seconds - (now - last)
                if remaining > 0:
                    raise NotificationTestError(
                        "notification_test_rate_limited",
                        f"测试发送过于频繁，请 {int(remaining) + 1} 秒后重试",
                        status_code=429,
                        retry_after=int(remaining) + 1,
                        category="rate_limited",
                    )

            self._last_sent[key] = now
            self._prune(now)

    async def reset(self) -> None:
        async with self._lock:
            self._last_sent.clear()

    def _prune(self, now: float) -> None:
        if len(self._last_sent) <= self.max_users:
            return
        for key in list(self._last_sent):
            if len(self._last_sent) <= self.max_users:
                break
            if now - self._last_sent[key] >= self.interval_seconds:
                self._last_sent.pop(key, None)


notification_test_rate_limiter = NotificationTestRateLimiter()


def _truncate(value: Any, limit: int) -> str:
    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    return text[:limit]


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class NotificationTestService:
    """读取规则绑定的渠道并发送一条真实测试消息。"""

    def __init__(
        self,
        db_manager: Any,
        *,
        sender: Optional[Any] = None,
        limiter: Optional[NotificationTestRateLimiter] = None,
    ) -> None:
        self.db_manager = db_manager
        self.sender = sender or NotificationSender()
        self.limiter = limiter

    async def send_rule_test(
        self,
        rule_id: int,
        user_id: int,
        user_info: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        try:
            normalized_rule_id = int(rule_id)
        except (TypeError, ValueError) as exc:
            raise NotificationTestError(
                "notification_rule_invalid",
                "通知规则不存在",
                status_code=404,
            ) from exc

        try:
            normalized_user_id = int(user_id)
        except (TypeError, ValueError) as exc:
            raise NotificationTestError(
                "notification_rule_invalid",
                "通知规则不存在",
                status_code=404,
            ) from exc

        target = self._load_target(normalized_rule_id, normalized_user_id)
        if not target:
            raise NotificationTestError(
                "notification_rule_not_found",
                "通知规则不存在或无权访问",
                status_code=404,
            )

        channel_type = normalize_channel_type(target.get("channel_type"))
        channel_name = str(target.get("channel_name") or channel_type or "通知渠道")

        if self.limiter is not None:
            await self.limiter.acquire(normalized_user_id)

        message = self._build_message(target, user_info)
        request_id = uuid.uuid4().hex
        started = time.perf_counter()

        try:
            receipt = await self.sender.send(
                channel_type,
                target.get("channel_config"),
                message,
                request_id=request_id,
            )
        except NotificationChannelConfigError as exc:
            raise NotificationTestError(
                "notification_config_invalid",
                str(exc) or "通知渠道配置不完整，请先补全渠道配置",
                status_code=400,
                category="config",
            ) from exc
        except NotificationSendTimeout as exc:
            raise NotificationTestError(
                exc.code,
                exc.public_message,
                status_code=504,
                category=exc.category,
            ) from exc
        except NotificationSendError as exc:
            raise NotificationTestError(
                exc.code,
                exc.public_message,
                status_code=502,
                category=exc.category,
            ) from exc
        except Exception as exc:  # noqa: BLE001 - 兜底不让底层异常直接冒泡
            logger.error(f"source=notification_test rule_id={normalized_rule_id} result=unexpected_error")
            raise NotificationTestError(
                "notification_send_failed",
                "通知渠道发送失败，请检查网络或服务配置",
                status_code=502,
                category="unknown",
            ) from exc

        duration_ms = int((time.perf_counter() - started) * 1000)
        sent_channel_type = getattr(receipt, "channel_type", None) or channel_type

        logger.info(
            f"source=notification_test rule_id={normalized_rule_id} "
            f"channel_type={sent_channel_type} result=sent duration_ms={duration_ms}"
        )

        return {
            "success": True,
            "message": "测试消息已发送",
            "request_id": request_id,
            "channel": {
                "id": int(target.get("channel_id") or 0),
                "name": channel_name,
                "type": sent_channel_type,
            },
            "sent_at": _now_iso(),
            "duration_ms": duration_ms,
        }

    def _load_target(self, rule_id: int, user_id: int) -> Optional[Dict[str, Any]]:
        try:
            return self.db_manager.get_notification_test_target(rule_id, user_id)
        except Exception as exc:  # noqa: BLE001 - 数据库异常不应带出内部细节
            logger.error(f"读取通知测试目标失败: {exc}")
            return None

    @staticmethod
    def _build_message(
        target: Dict[str, Any],
        user_info: Optional[Dict[str, Any]] = None,
    ) -> str:
        rule_name = _truncate(target.get("name"), MAX_RULE_NAME_CHARS)
        event_types = target.get("event_types") or []
        if isinstance(event_types, (list, tuple)):
            event_summary = "、".join(_truncate(item, 24) for item in event_types[:6]) or "未选择"
        else:
            event_summary = "未选择"

        username = _truncate((user_info or {}).get("username"), 40)
        lines = [
            "闲鱼超级管家 · 通知测试",
            f"规则：{rule_name or '（未命名规则）'}",
            f"订阅事件：{event_summary}",
            f"发送时间：{_now_iso()}",
            "",
            "若你收到这条消息，说明该通知渠道配置可用。",
        ]
        if username:
            lines.insert(3, f"操作账号：{username}")
        return "\n".join(lines)
