"""发货内容分段发送器。

`app/reply_server.py` 的批量发货接口依赖本模块的 :func:`send_payload`：
拿到一条卡券的发货内容后，把它发送给买家，并返回**实际发送出去的段数**
（接口会用它判断"是否全部发出"，进而决定要不要去平台确认发货）。

调用点契约（`reply_server.py` 第 23 行 / 第 8608 行）::

    from app.delivery_template import send_payload as send_delivery_payload
    segment_count = await send_delivery_payload(
        live_instance, live_instance.ws, chat_id, buyer_id, content
    )

分段规则：

* 图片标记行（``__IMAGE_SEND__[卡券ID|]图片地址``）单独成段，走图片消息通道
  （与 `XianyuAutoAsync._handle_auto_delivery` 里的既有格式保持一致）；
* 文本按换行聚合，累计长度超过 ``DELIVERY_SEGMENT_MAX_CHARS``（默认 500）时
  拆成多条消息；**没超上限的短内容不会被拆开**，仍然是一条消息；
* 段与段之间默认间隔 ``DELIVERY_SEGMENT_DELAY`` 秒（默认 0.6），降低被平台
  判定为机器刷屏的概率。

两个阈值都可以用环境变量覆盖，不需要改代码：
``DELIVERY_SEGMENT_MAX_CHARS``、``DELIVERY_SEGMENT_DELAY``。

说明：本文件在上游公开源码里缺失（作者漏提交），这里按调用点与仓库内既有的
发送实现（`XianyuAutoAsync.send_msg` / `send_image_msg`）对齐补齐，不引入新的
平台协议。
"""

from __future__ import annotations

import asyncio
import os
from typing import Any, List, Optional, Sequence, Tuple, Union

from loguru import logger

__all__ = [
    "IMAGE_MARKER",
    "parse_image_marker",
    "split_text_segments",
    "build_segments",
    "send_payload",
]

# 图片发送标记，格式由 XianyuAutoAsync 生成：__IMAGE_SEND__[card_id|]image_url
IMAGE_MARKER = "__IMAGE_SEND__"

DEFAULT_MAX_SEGMENT_CHARS = 500
DEFAULT_SEGMENT_DELAY = 0.6

# 段类型：文本段为 ("text", str)，图片段为 ("image", (card_id, url))
Segment = Tuple[str, Any]


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None or str(raw).strip() == "":
        return default
    try:
        value = int(str(raw).strip())
    except (TypeError, ValueError):
        logger.warning(f"[发货分段] 环境变量 {name}={raw!r} 不是整数，回退默认值 {default}")
        return default
    if value <= 0:
        logger.warning(f"[发货分段] 环境变量 {name}={value} 必须为正整数，回退默认值 {default}")
        return default
    return value


def _env_float(name: str, default: float) -> float:
    raw = os.getenv(name)
    if raw is None or str(raw).strip() == "":
        return default
    try:
        value = float(str(raw).strip())
    except (TypeError, ValueError):
        logger.warning(f"[发货分段] 环境变量 {name}={raw!r} 不是数字，回退默认值 {default}")
        return default
    if value < 0:
        logger.warning(f"[发货分段] 环境变量 {name}={value} 不能为负数，回退默认值 {default}")
        return default
    return value


def parse_image_marker(value: str) -> Tuple[Optional[int], Optional[str]]:
    """解析 ``__IMAGE_SEND__[card_id|]image_url``，返回 ``(card_id, image_url)``。

    兼容旧格式（没有卡券 ID 前缀）。解析不出图片地址时返回 ``(None, None)``。
    """
    if not value or not value.startswith(IMAGE_MARKER):
        return None, None

    payload = value[len(IMAGE_MARKER):].strip()
    if not payload:
        return None, None

    card_id: Optional[int] = None
    image_ref = payload
    if "|" in payload:
        card_id_str, rest = payload.split("|", 1)
        image_ref = rest.strip()
        try:
            card_id = int(card_id_str.strip())
        except (TypeError, ValueError):
            logger.warning(f"[发货分段] 图片标记里的卡券 ID 非法，按无卡券 ID 处理: {card_id_str!r}")
            card_id = None

    return (card_id, image_ref) if image_ref else (card_id, None)


def split_text_segments(text: str, max_segment_chars: int = DEFAULT_MAX_SEGMENT_CHARS) -> List[str]:
    """把一段文本按换行聚合切分，单条不超过 ``max_segment_chars`` 个字符。

    尽量在换行处切分；单行本身就超长时按长度硬切，避免整条消息被平台拒收。
    """
    if not text:
        return []

    normalized = str(text).replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    if not normalized.strip():
        return []

    limit = max(1, int(max_segment_chars))
    segments: List[str] = []
    buffer = ""

    for raw_line in normalized.split("\n"):
        line = raw_line.rstrip()
        candidate = line if not buffer else f"{buffer}\n{line}"

        if len(candidate) <= limit:
            buffer = candidate
            continue

        if buffer:
            segments.append(buffer)
            buffer = ""

        if len(line) <= limit:
            buffer = line
            continue

        # 单行超长：先按长度硬切，最后一段留给后续行继续聚合
        while len(line) > limit:
            segments.append(line[:limit])
            line = line[limit:]
        buffer = line

    if buffer:
        segments.append(buffer)

    return [segment for segment in segments if segment.strip()]


def build_segments(
    content: Union[str, None],
    max_segment_chars: int = DEFAULT_MAX_SEGMENT_CHARS,
) -> List[Segment]:
    """把一条卡券的发货内容拆成有序的段列表（文本段与图片段可混排）。"""
    if content is None:
        return []

    normalized = str(content).replace("\r\n", "\n").replace("\r", "\n")
    segments: List[Segment] = []
    text_buffer: List[str] = []

    def flush_text() -> None:
        if not text_buffer:
            return
        chunk = "\n".join(text_buffer)
        text_buffer.clear()
        segments.extend(("text", part) for part in split_text_segments(chunk, max_segment_chars))

    for line in normalized.split("\n"):
        stripped = line.strip()
        if stripped.startswith(IMAGE_MARKER):
            flush_text()
            card_id, image_ref = parse_image_marker(stripped)
            if image_ref:
                segments.append(("image", (card_id, image_ref)))
            else:
                logger.warning("[发货分段] 图片标记缺少图片地址，已忽略该行")
            continue
        text_buffer.append(line)

    flush_text()
    return segments


async def send_payload(
    instance: Any,
    ws: Any,
    chat_id: str,
    buyer_id: str,
    content: Union[str, None],
    *,
    max_segment_chars: Optional[int] = None,
    segment_delay: Optional[float] = None,
) -> int:
    """把一条发货内容按段发送给买家，返回成功发送的段数。

    ``instance`` 是运行中的 ``XianyuLive`` 实例（需要提供 ``send_msg`` /
    ``send_image_msg``），``ws`` 为其 websocket 连接，``chat_id`` 是会话 ID，
    ``buyer_id`` 是买家 ID。任一段发送失败都会抛异常，由调用方记录并计入失败。
    """
    limit = max_segment_chars if max_segment_chars is not None else _env_int(
        "DELIVERY_SEGMENT_MAX_CHARS", DEFAULT_MAX_SEGMENT_CHARS
    )
    delay = segment_delay if segment_delay is not None else _env_float(
        "DELIVERY_SEGMENT_DELAY", DEFAULT_SEGMENT_DELAY
    )

    segments = build_segments(content, limit)
    if not segments:
        logger.warning("[发货分段] 发货内容为空，跳过发送")
        return 0

    target_ws = ws if ws is not None else getattr(instance, "ws", None)
    if target_ws is None:
        raise RuntimeError("WebSocket 连接不可用，无法发送发货内容")

    sent_count = 0
    for index, (kind, payload) in enumerate(segments):
        if index > 0 and delay > 0:
            await asyncio.sleep(delay)

        if kind == "image":
            card_id, image_ref = payload
            await instance.send_image_msg(
                target_ws, chat_id, buyer_id, image_ref, card_id=card_id
            )
            logger.info(
                f"[发货分段] 已发送图片段 {index + 1}/{len(segments)} "
                f"(买家={buyer_id}, 卡券ID={card_id})"
            )
        else:
            await instance.send_msg(target_ws, chat_id, buyer_id, payload)
            logger.info(
                f"[发货分段] 已发送文本段 {index + 1}/{len(segments)} "
                f"(买家={buyer_id}, 长度={len(payload)})"
            )

        sent_count += 1

    return sent_count
