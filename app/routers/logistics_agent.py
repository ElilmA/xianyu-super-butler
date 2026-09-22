"""物流报价 Agent 的配置与试算路由。

`app/reply_server.py` 第 40 行导入 :func:`create_logistics_agent_router`，
第 363 行挂载到应用。数据库里已有的 ``logistics_agent_settings`` /
``logistics_quote_sessions`` / ``logistics_quote_send_logs`` /
``logistics_agent_training_*`` 说明原作者的完整形态是"私聊里自动询价并报价"，
但这部分业务层（含模型调用与话术模板）没有随公开源码发布，无法复原其行为。

因此本模块只补齐**可验证、可自洽**的部分，不假装实现 AI 自动报价：

* ``GET/PUT /api/logistics/agent/settings``：按账号读取/保存配置
  （启用开关、模型名、绑定报价表、推荐策略、无线路策略、商品范围、体积重系数、
  运费模板），字段与 ``logistics_agent_settings`` 表一一对应；
* ``GET /api/logistics/agent/quote-preview``：用已导入的线路做一次真实试算，
  供配置排查与人工报价时直接复制价格；
* ``GET/POST/DELETE /api/logistics/agent/training-rounds``：保存/读取人工构造的
  询价对话样本（结构化的 messages），后续接入模型时可作为训练/回归数据。

自动回复侧若要真正"自动报价"，需要另外实现"买家消息 → 提取地址重量 → 调用
本模块的试算 → 生成话术 → 发送并登记 send_logs"的链路，本文件不擅自插入
消息处理流程，避免影响现有自动回复与自动发货。
"""

from __future__ import annotations

import json
import math
from typing import Any, Callable, Dict, List, Optional

from fastapi import APIRouter, Body, Depends, HTTPException, Query
from loguru import logger
from pydantic import BaseModel, Field

__all__ = [
    "create_logistics_agent_router",
    "calculate_price",
    "match_routes",
    "VOLUMETRIC_DIVISOR_DEFAULT",
]

VOLUMETRIC_DIVISOR_DEFAULT = 6000.0
MONEY_DIGITS = 2

RECOMMEND_MODES = {"lowest", "fastest", "balanced"}
NO_ROUTE_POLICIES = {"manual", "fallback", "reject"}
ITEM_SCOPES = {"all", "selected"}


def _round_money(value: float) -> float:
    return round(float(value) + 1e-9, MONEY_DIGITS)


def _city_key(value: Any) -> str:
    text = str(value or "").strip()
    for suffix in ("特别行政区", "自治州", "地区", "自治区", "省", "市", "盟", "县", "区"):
        if text.endswith(suffix) and len(text) > len(suffix):
            text = text[: -len(suffix)]
            break
    return text


def _is_num(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def calculate_price(route: Dict[str, Any], weight_kg: float) -> Dict[str, Any]:
    """按线路的计价模型算价，返回(价格, 说明, 是否命中)。

    ``route`` 是 ``logistics_quote_routes.price_model`` 解析后的字典，
    字段与报价表解析结果保持一致。报价表里写明的"最低 N 元"会作为价格下限。
    """
    weight = float(weight_kg)
    if weight <= 0 or not math.isfinite(weight):
        return {"price": None, "detail": "重量必须大于 0", "matched": False}

    model = route.get("price_model") if isinstance(route, dict) else None
    if not isinstance(model, dict):
        model = {}

    result = _calculate_price_without_floor(model, weight)
    min_price = model.get("min_price")
    if result.get("price") is not None and _is_num(min_price) and result["price"] < float(min_price):
        result = {
            "price": _round_money(min_price),
            "detail": f"{result['detail']}；按最低计费 {float(min_price):g} 元",
            "matched": True,
        }
    return result


def _calculate_price_without_floor(model: Dict[str, Any], weight: float) -> Dict[str, Any]:
    """核心计价逻辑（不含最低计费兜底）。"""

    rule_type = model.get("rule_type")
    first_weight = model.get("first_weight_kg")
    first_price = model.get("first_price")
    continued_unit = model.get("continued_unit_kg") or 1.0
    continued_price = model.get("continued_price")
    continued_tiers = model.get("continued_tiers")
    fixed_tiers = model.get("fixed_tiers")

    if rule_type == "fixed_tiers" or (fixed_tiers and rule_type in (None, "fixed_tiers_overflow")):
        tiers = [tier for tier in (fixed_tiers or []) if _is_num(tier.get("up_to_kg")) and _is_num(tier.get("price"))]
        tiers.sort(key=lambda tier: float(tier["up_to_kg"]))
        for tier in tiers:
            if weight <= float(tier["up_to_kg"]):
                return {
                    "price": _round_money(tier["price"]),
                    "detail": f"{tier['up_to_kg']:g}kg 以内固定 {tier['price']:g} 元",
                    "matched": True,
                }
        if tiers and rule_type == "fixed_tiers":
            return {
                "price": None,
                "detail": f"超出最大档位（{tiers[-1]['up_to_kg']:g}kg），需按续重另行询价",
                "matched": False,
            }

    if rule_type == "fixed_tiers_overflow" and fixed_tiers:
        base = fixed_tiers[0]
        base_price = float(base.get("price") or 0)
        base_kg = float(base.get("up_to_kg") or 0)
        extra_price = continued_price if _is_num(continued_price) else None
        if extra_price is None:
            return {"price": _round_money(base_price), "detail": "固定价，未提供超重价格", "matched": True}
        extra_weight = max(0.0, weight - base_kg)
        return {
            "price": _round_money(base_price + extra_weight * extra_price),
            "detail": (
                f"{base_kg:g}kg 内 {base_price:g} 元，超出按 {extra_price:g} 元/kg"
                f"（超出 {extra_weight:g}kg）"
            ),
            "matched": True,
        }

    if rule_type == "banded_additional" and continued_tiers:
        tiers = [
            tier
            for tier in continued_tiers
            if _is_num(tier.get("min_exclusive_kg")) and _is_num(tier.get("price_per_kg"))
        ]
        tiers.sort(key=lambda tier: float(tier["min_exclusive_kg"]))
        # 取"起点低于当前重量"里起点最高的那一档，天然兼容开放式末档
        applicable = [tier for tier in tiers if weight > float(tier["min_exclusive_kg"])]
        selected = applicable[-1] if applicable else None
        base_kg = float(first_weight) if _is_num(first_weight) else None
        base_price = float(first_price) if _is_num(first_price) else 0.0

        if not applicable:
            # 重量落在首重区间内：按首重价计费
            if _is_num(first_price) and (base_kg is None or weight <= base_kg):
                return {
                    "price": _round_money(base_price),
                    "detail": f"首重 {base_kg:g}kg 内 {base_price:g} 元" if base_kg else f"最低 {base_price:g} 元",
                    "matched": True,
                }
            if tiers:
                return {
                    "price": None,
                    "detail": f"重量低于最小分档起点（{tiers[0]['min_exclusive_kg']:g}kg）",
                    "matched": False,
                }
        if selected is not None:
            rate = float(selected["price_per_kg"])
            low = float(selected["min_exclusive_kg"])
            basis = selected.get("basis") or ("continued" if _is_num(first_weight) else "total")
            if basis == "continued":
                billable = max(0.0, weight - low)
            else:
                billable = weight
            total = billable * rate + base_price
            detail = f"{low:g}kg 以上 {rate:g} 元/kg × {billable:g}kg"
            if basis == "continued":
                detail += "（续重口径）"
            if base_price:
                detail += f" + 首重 {base_price:g} 元"
            return {"price": _round_money(total), "detail": detail, "matched": True}

    if rule_type == "minimum_then_per_kg" and continued_tiers:
        tiers = [tier for tier in continued_tiers if _is_num(tier.get("price_per_kg"))]
        tiers.sort(key=lambda tier: float(tier.get("min_exclusive_kg") or 0))
        rate = float(tiers[-1]["price_per_kg"]) if tiers else None
        if rate is not None:
            minimum_kg = float(tiers[0].get("min_exclusive_kg") or 0)
            billable = max(weight, minimum_kg)
            return {
                "price": _round_money(billable * rate),
                "detail": f"{rate:g} 元/kg，最低按 {minimum_kg:g}kg 计（计费重 {billable:g}kg）",
                "matched": True,
            }

    if _is_num(first_price) or _is_num(continued_price):
        base_kg = float(first_weight) if _is_num(first_weight) else 1.0
        base_price = float(first_price) if _is_num(first_price) else 0.0
        extra_rate = float(continued_price) if _is_num(continued_price) else 0.0
        step = float(continued_unit) if _is_num(continued_unit) and float(continued_unit) > 0 else 1.0
        if weight <= base_kg:
            return {
                "price": _round_money(base_price),
                "detail": f"首重 {base_kg:g}kg 内 {base_price:g} 元",
                "matched": True,
            }
        extra_weight = weight - base_kg
        extra_units = math.ceil(extra_weight / step - 1e-9)
        total = base_price + extra_units * step * extra_rate
        return {
            "price": _round_money(total),
            "detail": (
                f"首重 {base_kg:g}kg {base_price:g} 元 + 续重 {extra_units:g}×{step:g}kg "
                f"× {extra_rate:g} 元"
            ),
            "matched": True,
        }

    if _is_num(model.get("quote")):
        return {
            "price": _round_money(model["quote"]),
            "detail": f"整条线路一口价 {model['quote']:g} 元（未区分重量）",
            "matched": True,
        }

    return {"price": None, "detail": "该线路的价格模型无法识别，请人工复核报价表", "matched": False}


def match_routes(
    routes: List[Dict[str, Any]],
    *,
    destination_city: str,
    origin_city: Optional[str] = None,
    carrier: Optional[str] = None,
    book_kind: Optional[str] = None,
    limit: int = 20,
) -> List[Dict[str, Any]]:
    """按目的地（可选始发地/承运商/报价表类型）筛选候选线路。"""
    destination_key = _city_key(destination_city)
    if not destination_key:
        return []
    origin_key = _city_key(origin_city) if origin_city else ""
    carrier_key = str(carrier or "").strip()

    matched: List[Dict[str, Any]] = []
    for route in routes:
        if not isinstance(route, dict):
            continue
        if book_kind and route.get("book_kind") and route["book_kind"] != book_kind:
            continue
        if carrier_key and carrier_key not in str(route.get("carrier") or ""):
            continue
        dest_key = _city_key(route.get("dest_city")) or _city_key(route.get("dest_province"))
        if not dest_key:
            continue
        if dest_key not in destination_key and destination_key not in dest_key:
            continue
        if origin_key:
            route_origin = _city_key(route.get("origin_city"))
            if route_origin and route_origin not in origin_key and origin_key not in route_origin:
                continue
        matched.append(route)

    return matched[: max(1, min(int(limit), 100))]


class LogisticsAgentSettingsUpdate(BaseModel):
    enabled: Optional[bool] = None
    model_name: Optional[str] = Field(default=None, max_length=120)
    book_ids: Optional[List[int]] = None
    auto_send: Optional[bool] = None
    recommend_mode: Optional[str] = None
    no_route_policy: Optional[str] = None
    item_scope: Optional[str] = None
    item_ids: Optional[List[str]] = None
    carrier_config: Optional[Dict[str, Any]] = None
    default_volume_ratios: Optional[Dict[str, Any]] = None
    pricing_config: Optional[Dict[str, Any]] = None
    templates: Optional[Dict[str, Any]] = None


class TrainingMessage(BaseModel):
    role: str = Field(pattern="^(buyer|agent|system)$")
    content: str = Field(min_length=1, max_length=4000)
    position: Optional[int] = None
    decision: Optional[Dict[str, Any]] = None


class TrainingRoundCreate(BaseModel):
    cookie_id: str = Field(min_length=1, max_length=120)
    thread_id: str = Field(min_length=1, max_length=120)
    name: str = Field(default="", max_length=200)
    round_id: Optional[str] = Field(default=None, max_length=120)
    messages: List[TrainingMessage] = Field(default_factory=list)


def create_logistics_agent_router(
    get_current_user: Callable[..., Dict[str, Any]],
    db_manager: Any,
) -> APIRouter:
    router = APIRouter(prefix="/api/logistics/agent", tags=["logistics-agent"])

    def _user_id(current_user: Dict[str, Any]) -> int:
        return int(current_user["user_id"])

    def _require_account(cookie_id: str, current_user: Dict[str, Any]) -> Dict[str, Any]:
        details = db_manager.get_cookie_details(cookie_id)
        if not details or details.get("user_id") != current_user["user_id"]:
            raise HTTPException(status_code=404, detail="账号不存在或无权限")
        return details

    def _load_json(value: Any, default: Any) -> Any:
        if value is None or value == "":
            return default
        if isinstance(value, (dict, list)):
            return value
        if isinstance(value, (bytes, bytearray)):
            value = value.decode("utf-8", errors="replace")
        if isinstance(value, str):
            try:
                return json.loads(value)
            except (TypeError, ValueError):
                return default
        return default

    def _settings_payload(row: Optional[Dict[str, Any]], cookie_id: str) -> Dict[str, Any]:
        row = row or {}
        return {
            "cookie_id": cookie_id,
            "enabled": bool(row.get("enabled")),
            "model_name": row.get("model_name") or "deepseek-v4-flash",
            "book_ids": _load_json(row.get("book_ids"), []),
            "auto_send": bool(row.get("auto_send")),
            "recommend_mode": row.get("recommend_mode") or "lowest",
            "no_route_policy": row.get("no_route_policy") or "manual",
            "item_scope": row.get("item_scope") or "all",
            "item_ids": _load_json(row.get("item_ids"), []),
            "carrier_config": _load_json(row.get("carrier_config"), {}),
            "default_volume_ratios": _load_json(row.get("default_volume_ratios"), {}),
            "pricing_config": _load_json(row.get("pricing_config"), {}),
            "templates": _load_json(row.get("templates"), {}),
            "created_at": row.get("created_at"),
            "updated_at": row.get("updated_at"),
        }

    @router.get("/settings")
    def get_agent_settings(
        cookie_id: str = Query(..., min_length=1),
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        _require_account(cookie_id, current_user)
        row = db_manager.get_logistics_agent_settings(cookie_id)
        return {"success": True, "settings": _settings_payload(row, cookie_id)}

    @router.put("/settings")
    def update_agent_settings(
        cookie_id: str = Query(..., min_length=1),
        payload: LogisticsAgentSettingsUpdate = Body(...),
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        _require_account(cookie_id, current_user)

        if payload.recommend_mode is not None and payload.recommend_mode not in RECOMMEND_MODES:
            raise HTTPException(status_code=400, detail="推荐策略仅支持 lowest / fastest / balanced")
        if payload.no_route_policy is not None and payload.no_route_policy not in NO_ROUTE_POLICIES:
            raise HTTPException(status_code=400, detail="无线路策略仅支持 manual / fallback / reject")
        if payload.item_scope is not None and payload.item_scope not in ITEM_SCOPES:
            raise HTTPException(status_code=400, detail="商品范围仅支持 all / selected")

        values: Dict[str, Any] = {}
        simple_fields = {
            "enabled": lambda value: 1 if value else 0,
            "model_name": lambda value: str(value)[:120],
            "auto_send": lambda value: 1 if value else 0,
            "recommend_mode": str,
            "no_route_policy": str,
            "item_scope": str,
        }
        json_fields = {
            "book_ids": list,
            "item_ids": list,
            "carrier_config": dict,
            "default_volume_ratios": dict,
            "pricing_config": dict,
            "templates": dict,
        }
        provided = payload.model_dump(exclude_unset=True, exclude_none=True)
        for field, caster in simple_fields.items():
            if field in provided:
                values[field] = caster(provided[field])
        for field in json_fields:
            if field in provided:
                values[field] = json.dumps(provided[field], ensure_ascii=False)

        try:
            row = db_manager.upsert_logistics_agent_settings(cookie_id, values)
        except Exception as exc:  # noqa: BLE001 - 不暴露底层异常
            logger.error(f"保存物流 Agent 配置失败: {exc}")
            raise HTTPException(status_code=500, detail="保存失败，请稍后重试") from exc

        return {"success": True, "settings": _settings_payload(row, cookie_id)}

    @router.get("/quote-books")
    def list_agent_quote_books(current_user: Dict[str, Any] = Depends(get_current_user)):
        """Agent 可绑定的报价表清单（只返回摘要）。"""
        books = db_manager.list_logistics_quote_books(_user_id(current_user))
        return {"success": True, "books": books}

    @router.get("/quote-preview")
    def quote_preview(
        destination_city: str = Query(..., min_length=1),
        weight_kg: float = Query(..., gt=0, le=10000),
        origin_city: Optional[str] = None,
        carrier: Optional[str] = None,
        book_kind: Optional[str] = None,
        length_cm: Optional[float] = Query(default=None, gt=0),
        width_cm: Optional[float] = Query(default=None, gt=0),
        height_cm: Optional[float] = Query(default=None, gt=0),
        volume_divisor: float = Query(default=VOLUMETRIC_DIVISOR_DEFAULT, gt=0),
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        """用已导入线路做一次真实试算（不写库、不发送任何消息）。"""
        billed_weight = float(weight_kg)
        volume_weight = None
        if length_cm and width_cm and height_cm:
            volume_weight = (float(length_cm) * float(width_cm) * float(height_cm)) / float(volume_divisor)
            billed_weight = max(billed_weight, volume_weight)

        routes = db_manager.list_logistics_quote_routes(
            _user_id(current_user), book_kind=book_kind
        )
        if not routes:
            return {
                "success": True,
                "matched": 0,
                "options": [],
                "warning": "尚未导入任何报价线路，请先上传报价表",
            }

        candidates = match_routes(
            routes,
            destination_city=destination_city,
            origin_city=origin_city,
            carrier=carrier,
            book_kind=book_kind,
        )

        options: List[Dict[str, Any]] = []
        for route in candidates:
            result = calculate_price(route, billed_weight)
            if result["price"] is None:
                continue
            options.append(
                {
                    "route_id": route.get("id"),
                    "carrier": route.get("carrier"),
                    "book_kind": route.get("book_kind"),
                    "origin_city": route.get("origin_city"),
                    "destination_city": route.get("dest_city"),
                    "price": result["price"],
                    "detail": result["detail"],
                    "eta": (route.get("price_model") or {}).get("eta"),
                }
            )

        options.sort(key=lambda item: item["price"])
        cheapest = options[0] if options else None
        return {
            "success": True,
            "matched": len(options),
            "weight_kg": round(float(weight_kg), 3),
            "volume_weight_kg": round(volume_weight, 3) if volume_weight is not None else None,
            "billed_weight_kg": round(billed_weight, 3),
            "options": options,
            "recommended": cheapest,
            "warning": "" if options else "没有可匹配且可计价的线路，建议人工确认目的地或补充报价表",
        }

    @router.get("/training-rounds")
    def list_training_rounds(
        cookie_id: Optional[str] = None,
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        rounds = db_manager.list_logistics_training_rounds(_user_id(current_user), cookie_id)
        return {"success": True, "rounds": rounds}

    @router.post("/training-rounds")
    def create_training_round(
        payload: TrainingRoundCreate,
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        _require_account(payload.cookie_id, current_user)
        messages = [
            {
                "role": message.role,
                "content": message.content,
                **({"position": message.position} if message.position is not None else {}),
                **({"decision": message.decision} if message.decision else {}),
            }
            for message in payload.messages
        ]
        round_id = db_manager.save_logistics_training_round(
            user_id=_user_id(current_user),
            cookie_id=payload.cookie_id,
            thread_id=payload.thread_id,
            name=payload.name,
            messages=messages,
            round_id=payload.round_id,
        )
        if not round_id:
            raise HTTPException(status_code=500, detail="保存训练样本失败，请稍后重试")
        return {"success": True, "id": round_id}

    @router.delete("/training-rounds/{round_id}")
    def delete_training_round(
        round_id: str,
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        if not db_manager.delete_logistics_training_round(round_id, _user_id(current_user)):
            raise HTTPException(status_code=404, detail="训练样本不存在")
        return {"success": True}

    return router
