"""物流报价表解析与管理路由。

`app/reply_server.py` 第 39 行导入 :func:`create_logistics_quote_router`，
第 360 行挂载到应用；前端 `frontend/services/api.ts` 的
``parseQuoteSource`` / ``listQuoteBooks`` / ``createQuoteBook`` / ``deleteQuoteBook``
依赖下面三个接口：

* ``POST /api/logistics/quote-sources/parse`` 上传报价表并返回解析结果（不落库）；
* ``GET/POST /api/logistics/quote-books`` 列出 / 保存报价表；
* ``DELETE /api/logistics/quote-books/{book_id}`` 删除报价表。

接口形状严格对齐前端 TypeScript 类型（``LogisticsQuoteParseResponse``、
``LogisticsQuoteBook``）。报价表是卖家自己整理的 Excel/CSV，列名和分档写法不统一，
因此解析采用"表头同义词匹配 + 分档文本解析"的启发式方案：

* 表头行通过同义词打分识别（首重/首重价/续重/每公斤/目的地/承运商…）；
* 块状表格（一行写承运商名，后续行为该承运商的线路）会拆成 carrier 块；
* 计价规则归纳为 ``first_additional`` / ``fixed_tiers`` / ``fixed_tiers_overflow``
  / ``banded_additional`` / ``minimum_then_per_kg`` 五类，无法识别的行标记
  ``review_state='review'`` 而不丢弃，交由人工复核。

本文件在上游公开源码里缺失（作者漏提交），这里按前端契约与数据库既有表结构补齐。
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from loguru import logger

try:  # pandas 在 requirements.txt 中，缺失时退化为纯 openpyxl 读取
    import pandas as _pd
except Exception:  # pragma: no cover - 依赖缺失时不影响其它接口
    _pd = None

SUPPORTED_EXTENSIONS = {".xlsx": "xlsx", ".xlsm": "xlsm", ".csv": "csv", ".txt": "csv"}
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
PARSER_VERSION = "xianyu-quote-parser/1.0"

# 表头同义词：键是内部字段名，值是按优先级排列的匹配片段
HEADER_ALIASES: Dict[str, Tuple[str, ...]] = {
    "carrier": ("承运商", "快递公司", "物流公司", "承运", "carrier", "渠道公司", "服务商"),
    "channel": ("渠道", "渠道名称", "线路名称", "产品", "产品名称", "服务名称", "channel"),
    "seller": ("卖家", "货主", "供应商", "seller"),
    "route": ("线路", "航线", "路线", "route"),
    "eta": ("时效", "时效说明", "预计时效", "运输时效", "eta", "参考时效"),
    "origin": ("起运地", "出发地", "始发地", "起运城市", "发货地", "origin"),
    "origin_province": ("起运省", "始发省", "发货省", "origin_province", "起运省份"),
    "origin_city": ("起运市", "始发市", "发货市", "origin_city", "起运城市名"),
    "destination": ("目的地", "目的地城市", "到达地", "收货地", "派送地", "destination"),
    "destination_province": ("目的省", "目的省份", "到达省", "destination_province"),
    "destination_city": ("目的市", "目的城市", "到达市", "destination_city"),
    "first_weight_kg": ("首重", "首重kg", "首重重量", "first_weight"),
    "first_price": ("首重价", "首重价格", "首重费用", "first_price"),
    "continued_unit_kg": ("续重单位", "续重步长", "计费单位", "每公斤", "continued_unit"),
    "continued_price": ("续重价", "续重价格", "续重费用", "续重单价", "continued_price"),
    "price_per_kg": ("每公斤价格", "单价", "公斤价", "price_per_kg", "元/kg", "元每公斤"),
    "quote": ("固定价", "一口价", "总价", "报价", "运费", "价格", "quote"),
}

# 计费规则关键词
FIXED_TIER_KEYWORDS = ("固定", "一口价", "包干", "统一价", "不分重量", "按票")
MINIMUM_KEYWORDS = ("最低", "起步", "起收")
BANDED_KEYWORDS = ("区间", "分档", "阶梯", "重量段")
CONTINUED_KEYWORDS = ("续重", "续重价", "每公斤", "公斤价", "元/kg", "元/千克")

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}


class QuoteParseError(ValueError):
    """上传文件无法解析时抛出，路由层翻译成 400。"""


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    if text.lower().strip() in {"nan", "none"}:
        return ""
    return re.sub(r"[\s\u3000()（）\[\]【】:：*　]+", "", text).lower()


def _matches_field_alias(field: str, value: Any) -> bool:
    """判断某个单元格文本是否真的符合该字段的同义词（用于复核表头映射）。"""
    normalized = _normalize_header(value)
    if not normalized:
        return False
    return any(
        _normalize_header(alias) and _normalize_header(alias) in normalized
        for alias in HEADER_ALIASES.get(field, ())
    )


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        if value.is_integer():
            return str(int(value))
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none", "nat"} else text


def _to_float(value: Any) -> Optional[float]:
    """宽松取数：单元格里只要有数字就取出来（重量、时效等字段用）。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return None if number != number else number
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    if not match:
        return None
    try:
        return float(match.group())
    except ValueError:
        return None


_CURRENCY_NOISE = re.compile(r"[¥￥元块\s]")
_UNIT_NOISE = re.compile(r"(?:/|每)\s*(?:kg|KG|公斤|千克)", re.IGNORECASE)


def _to_price(value: Any) -> Optional[float]:
    """严格取价：只有"纯金额"单元格才返回数值。

    含重量语义或分段语义的单元格（"5kg以内 20 元"、"1-5:3"）返回 None，
    交给 _parse_fixed_tiers / _parse_tiers 处理，避免把 "5" 当成价格。
    """
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        return None if number != number else number
    text = _cell_text(value)
    if not text:
        return None
    cleaned = _UNIT_NOISE.sub("", text)
    cleaned = _CURRENCY_NOISE.sub("", cleaned)
    if re.fullmatch(r"\d+(?:\.\d+)?", cleaned):
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


_NUMBER = r"-?\d+(?:\.\d+)?"
_RANGE_SEPARATOR = r"(?:-|~|—|–|至|到)"

_TIER_PATTERN = re.compile(
    rf"(?P<low>{_NUMBER})\s*(?:{_RANGE_SEPARATOR}\s*(?P<high>{_NUMBER})|(?P<open>\+)?)\s*"
    r"(?:kg|KG|Kg|公斤|千克|g)?\s*(?:以上|以下|以内|内)?\s*[:：,，]?\s*"
    r"(?:¥|￥)?\s*(?P<price>\d+(?:\.\d+)?)\s*(?:元|¥|￥)?\s*"
    r"(?:/|每)?\s*(?:kg|KG|Kg|公斤|千克)?"
    r"(?P<continued>续重|总重|首重)?"
)


def _parse_tiers(text: Any) -> Optional[List[Dict[str, Any]]]:
    """把 "1-5:3；5-10:2.5" 这类分档文本解析成 tier 列表。"""
    raw = _cell_text(text)
    if not raw:
        return None
    if _contains_min_price(raw):
        # 备注里的"最低50元"不是分档价
        return None
    tiers: List[Dict[str, Any]] = []
    for match in _TIER_PATTERN.finditer(raw):
        low = _to_float(match.group("low"))
        high = _to_float(match.group("high"))
        price = _to_float(match.group("price"))
        if price is None or low is None:
            continue
        tier: Dict[str, Any] = {
            "min_exclusive_kg": low,
            "max_inclusive_kg": high,
            "price_per_kg": price,
        }
        label = match.group("continued")
        if label == "续重":
            tier["basis"] = "continued"
        elif label == "总重":
            tier["basis"] = "total"
        elif high is None or match.group("open"):
            # "5kg以上 2元" 这种开放区间通常按每公斤计
            tier["basis"] = "total"
        tiers.append(tier)
    # 只保留价格单调不升或单调不降的连续分档，避免"1-5:3; 8-10:2.5"这类乱序被误读
    if len(tiers) > 1:
        tiers.sort(key=lambda item: item["min_exclusive_kg"])
    return tiers or None


def _contains_min_price(text: str) -> bool:
    """单元格是否只是"最低 N 元"这类起价说明（没有重量区间语义）。"""
    if not any(keyword in text for keyword in MINIMUM_KEYWORDS):
        return False
    if re.search(r"\d\s*(?:-|~|—|–|至|到)\s*\d", text):
        return False
    return bool(re.search(r"[¥￥]?\s*\d+(?:\.\d+)?\s*(?:元|块)?", text))


_MIN_PRICE_PATTERN = re.compile(r"(?:最低|起步|起收)\s*(?:价)?\s*[:：]?\s*[¥￥]?\s*(\d+(?:\.\d+)?)")


def _parse_min_price(text: Any) -> Optional[float]:
    """从备注文本里取出"最低 50 元"这类起价。"""
    raw = _cell_text(text)
    if not raw:
        return None
    match = _MIN_PRICE_PATTERN.search(raw)
    if not match:
        return None
    return _to_float(match.group(1))


_FIXED_TIER_PATTERN = re.compile(
    r"(?:≤|<=|不超过|小于等于|最大)?\s*(?P<upto>\d+(?:\.\d+)?)\s*(?:kg|KG|公斤|千克)\s*"
    r"(?:以内|以下|内)?\s*(?:¥|￥)?\s*(?P<price>\d+(?:\.\d+)?)"
)


def _parse_fixed_tiers(text: Any) -> Optional[List[Dict[str, Any]]]:
    """把 "5kg以内 20 元，10kg以内 35 元" 解析成固定档价格。"""
    raw = _cell_text(text)
    if not raw:
        return None
    tiers: List[Dict[str, Any]] = []
    for match in _FIXED_TIER_PATTERN.finditer(raw):
        upto = _to_float(match.group("upto"))
        price = _to_float(match.group("price"))
        if upto is None or price is None:
            continue
        tiers.append({"up_to_kg": upto, "price": price, "up_to": True})
    return tiers or None


def _split_origin_destination(text: str) -> Tuple[Optional[str], Optional[str]]:
    for separator in ("-", "—", "~", "→", "->", "到", "至"):
        if separator in text:
            left, _, right = text.partition(separator)
            left, right = left.strip(), right.strip()
            if left and right:
                return left, right
    return None, None


def _split_location(text: Any) -> Tuple[str, str, str]:
    """把 "浙江省金华市义乌市" 拆成 (省, 市, 原文)。认不出层级时全部塞进市。"""
    raw = _cell_text(text)
    if not raw:
        return "", "", ""
    province = ""
    city = ""
    province_match = re.match(r"^(.{2,8}省|.{2,8}自治区|.{2,8}特别行政区|北京市|上海市|天津市|重庆市)", raw)
    remainder = raw
    if province_match:
        province = province_match.group(1)
        remainder = raw[len(province):]
    city_match = re.match(r"^(.{1,8}市|.{1,8}自治州|.{1,8}地区|.{1,8}盟)", remainder)
    if city_match:
        city = city_match.group(1)
    elif remainder:
        city = remainder
    if not city:
        city = raw
    return province, city, raw


def _match_header_columns(row: List[str]) -> Dict[str, int]:
    """把一行的单元格按同义词一一映射到字段；列和字段都不会重复占用。"""
    candidates: List[Tuple[int, int, str]] = []
    for index, cell in enumerate(row):
        normalized = _normalize_header(cell)
        if not normalized:
            continue
        for field, aliases in HEADER_ALIASES.items():
            for alias in aliases:
                alias_normalized = _normalize_header(alias)
                if alias_normalized and alias_normalized in normalized:
                    # 排序键：别名越长越具体，越靠前的列优先
                    candidates.append((-len(alias_normalized), index, field))

    candidates.sort()
    mapping: Dict[str, int] = {}
    used_columns: set = set()
    for _, column, field in candidates:
        if field in mapping or column in used_columns:
            continue
        mapping[field] = column
        used_columns.add(column)
    return mapping


_DATA_CELL_PATTERN = re.compile(r"[\d¥￥]|\d\s*(?:-|~|至|到)\s*\d")


def _looks_like_data_cell(value: Any) -> bool:
    """含数字/金额/区间的单元格更可能是数据行而不是表头。"""
    text = _cell_text(value)
    return bool(text) and bool(_DATA_CELL_PATTERN.search(text))


def _is_header_row(mapping: Dict[str, int]) -> bool:
    if len(mapping) < 2:
        return False
    core = {"destination", "destination_city", "route", "carrier", "channel"}
    price = {"first_price", "continued_price", "price_per_kg", "quote", "fixed_tiers"}
    return bool(core & set(mapping)) and bool(price & set(mapping) or {"first_weight_kg"} & set(mapping))


def _guess_book_kind(mapping: Dict[str, int], sheet_name: str) -> Optional[str]:
    joined = " ".join([sheet_name or ""] + list(mapping.keys()))
    if any(keyword in joined for keyword in ("快递", "express", "首重", "续重")):
        return "express"
    if any(keyword in joined for keyword in ("物流", "专线", "logistics", "渠道")):
        return "logistics"
    return "express" if {"first_weight_kg", "continued_price"} & set(mapping) else None


def _infer_rule(
    row: Dict[str, Any],
    raw_rule_text: str,
) -> Tuple[Optional[str], List[str], float]:
    """推断计费规则类型，返回 (rule_type, issues, confidence)。"""
    issues: List[str] = []
    confidence = 1.0

    fixed_tiers = row.get("fixed_tiers")
    continued_tiers = row.get("continued_tiers")
    first_weight = row.get("first_weight_kg")
    first_price = row.get("first_price")
    continued_price = row.get("continued_price")
    quote = row.get("quote")

    if fixed_tiers:
        rule_type = "fixed_tiers"
        if len(fixed_tiers) > 1:
            rule_type = "fixed_tiers_overflow"
    elif continued_tiers and first_price is not None:
        rule_type = "banded_additional"
    elif continued_tiers and first_price is None:
        rule_type = "minimum_then_per_kg"
        confidence -= 0.15
    elif first_price is not None and continued_price is not None:
        rule_type = "first_additional"
    elif first_price is not None and first_weight is not None:
        rule_type = "first_additional"
        issues.append("缺少续重价格，超出首重后无法计费")
        confidence -= 0.35
    elif quote is not None:
        rule_type = "fixed_tiers_overflow"
        row.setdefault("fixed_tiers", [{"up_to_kg": 0, "price": quote, "up_to": False}])
        issues.append("仅识别到单一报价，未说明重量分档")
        confidence -= 0.25
    else:
        rule_type = None
        issues.append("未识别到计费规则")
        confidence -= 0.6

    if raw_rule_text and any(keyword in raw_rule_text for keyword in MINIMUM_KEYWORDS):
        confidence -= 0.05
    if row.get("destination_city") in (None, ""):
        issues.append("缺少目的地城市，无法参与地址匹配")
        confidence -= 0.3

    return rule_type, issues, max(0.0, min(1.0, confidence))


def _review_state(issues: List[str], confidence: float) -> str:
    if not issues and confidence >= 0.7:
        return "valid"
    if confidence <= 0.2:
        return "rejected"
    return "review"


def _read_tabular(data: bytes, extension: str, filename: str) -> List[Tuple[str, List[List[str]]]]:
    """读取上传文件，返回 [(sheet_name, 二维单元格文本)]。"""
    if extension == "csv":
        text = None
        for encoding in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                text = data.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise QuoteParseError("CSV 编码无法识别，请另存为 UTF-8 或 GBK 后重试")
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
            rows = list(csv.reader(io.StringIO(text), dialect))
        except Exception:  # noqa: BLE001 - Sniffer 失败时退回逗号分隔
            rows = list(csv.reader(io.StringIO(text)))
        return [("CSV", [[_cell_text(cell) for cell in row] for row in rows])]

    if _pd is None:
        raise QuoteParseError("服务器缺少 Excel 解析组件（pandas/openpyxl），无法解析该文件")

    try:
        with _pd.ExcelFile(io.BytesIO(data)) as workbook:
            sheets: List[Tuple[str, List[List[str]]]] = []
            for sheet_name in workbook.sheet_names:
                frame = workbook.parse(sheet_name, header=None, dtype=object)
                rows = [[_cell_text(value) for value in row] for row in frame.values.tolist()]
                sheets.append((str(sheet_name), rows))
            return sheets
    except QuoteParseError:
        raise
    except Exception as exc:  # noqa: BLE001 - 统一转成可读错误
        raise QuoteParseError(f"无法读取 {filename}：{exc}") from exc


def _detect_columns(sheet_name: str, rows: List[List[str]]) -> Tuple[Dict[str, int], int]:
    """在前 30 行里找表头，返回 (列映射, 表头行号)；找不到时返回宽表假设。"""
    best_mapping: Dict[str, int] = {}
    best_index = -1
    best_score = 0
    for index, row in enumerate(rows[:30]):
        mapping = _match_header_columns(row)
        if mapping and _looks_like_data_cell(row[mapping.get("destination", 0)] if row else ""):
            # "金华市-杭州市,25,次日达"这类数据行会被坐标号误判成表头，这里直接丢弃
            mapping = {}
        score = len(mapping)
        if score > best_score:
            best_mapping, best_index, best_score = mapping, index, score

    if _is_header_row(best_mapping):
        return best_mapping, best_index

    # 没有表头：按"线路 + 价格"的宽表假设列 0/1/2
    fallback = {"destination": 0, "quote": 1, "eta": 2}
    return fallback, -1


def _is_carrier_block_row(cells: List[str], mapping: Dict[str, int]) -> Optional[str]:
    """识别"整行只有承运商/服务商名字"的块状表头行。

    块状报价表常见写法：第一行只写一个服务商名字（"顺丰"、"中通-标快"），
    下面若干行是它的线路报价；名字列通常没有表头，因此没有映射到 carrier 列。

    判定条件：整行只有左侧一列有值、该值不像任何已知表头文字，
    也不是已映射列的取值（金额、城市名等）。
    """
    filled = [
        (index, _cell_text(value))
        for index, value in enumerate(cells[:12])
        if _cell_text(value)
    ]
    if len(filled) != 1:
        return None

    column, value = filled[0]
    carrier_column = mapping.get("carrier")
    is_carrier_column = carrier_column is not None and column == carrier_column
    # 必须是左侧列（或明确的承运商列），否则更可能是单字段残缺行
    if not is_carrier_column and column > 1:
        return None
    if len(value) > 40 or len(value) < 2:
        return None
    for field in set(mapping) | {"carrier"}:
        if _matches_field_alias(field, value):
            return None
    return value


def _remember_carrier(
    carriers: Dict[str, Dict[str, Any]],
    name: str,
    sheet_name: str,
    source: str,
    carrier_column_name: Optional[str],
    source_row: int,
) -> None:
    if not name:
        return
    entry = carriers.get(name)
    if entry is None:
        carriers[name] = {
            "name": name,
            "sheets": [sheet_name],
            "source": source,
            "carrier_column": carrier_column_name if source == "carrier_column" else None,
            "first_source_row": source_row,
        }
        return
    if sheet_name not in entry["sheets"]:
        entry["sheets"].append(sheet_name)
    if entry["source"] != source:
        entry["source"] = "mixed"


def _build_parse_rows(
    sheets: List[Tuple[str, List[List[str]]]],
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[str], List[str]]:
    """返回 (rows, carriers, services, warnings)。"""
    parsed_rows: List[Dict[str, Any]] = []
    services: List[Dict[str, Any]] = []
    warnings: List[str] = []
    carriers: Dict[str, Dict[str, Any]] = {}
    row_sequence = 0

    def _sheet_carrier(sheet_name: str) -> Optional[str]:
        """工作表名可以当承运商名，但 "Sheet1"/"CSV" 这类占位名不算。"""
        text = str(sheet_name or "").strip()
        if not text:
            return None
        normalized = text.lower()
        if normalized in {"csv", "txt"} or re.fullmatch(r"sheet\s*\d*", normalized):
            return None
        return text

    for sheet_name, rows in sheets:
        if not rows:
            continue
        mapping, header_index = _detect_columns(sheet_name, rows)
        if not mapping:
            warnings.append(f"工作表「{sheet_name}」没有识别到可用列，已跳过")
            continue

        carrier_column = mapping.get("carrier")
        carrier_column_name = (
            rows[header_index][carrier_column]
            if header_index >= 0 and carrier_column is not None
            and carrier_column < len(rows[header_index])
            and _matches_field_alias("carrier", rows[header_index][carrier_column])
            else None
        )
        header_cells = (
            rows[header_index] + [""] * (max(mapping.values()) + 1 - len(rows[header_index]))
            if header_index >= 0
            else []
        )
        recognized_columns = set(mapping.values())

        # 一个"服务块"= 一个承运商/服务名 + 它名下的线路行
        services_found: List[Dict[str, Any]] = []
        reported_carriers: set = set()
        block_name: Optional[str] = None
        block_rows = 0
        block_routes = 0
        block_rule_types: List[str] = []

        def _close_block() -> None:
            nonlocal block_rows, block_routes, block_rule_types
            if not block_rows:
                return
            services_found.append(
                {
                    "name": block_name or sheet_name,
                    "sheet_name": sheet_name,
                    "row_count": block_rows,
                    "route_count": block_routes,
                    "rule_type": block_rule_types[0] if block_rule_types else "unknown",
                    "book_kind": _guess_book_kind(mapping, sheet_name),
                    "mapping": {field: str(column) for field, column in mapping.items()},
                }
            )
            block_rows = 0
            block_routes = 0
            block_rule_types = []

        def _is_repeated_header_row(row: List[str]) -> bool:
            """块状表里每个承运商块都会重复一遍表头，需要跳过。"""
            if not header_cells or not row:
                return False
            checked = 0
            for field, column in mapping.items():
                if column >= len(row) or column >= len(header_cells):
                    continue
                header_text = _cell_text(header_cells[column])
                if not header_text:
                    continue
                checked += 1
                if _cell_text(row[column]) != header_text:
                    return False
            return checked >= 2

        for index, row in enumerate(rows):
            if index <= header_index:
                continue
            # 右侧列会随行递减，先补齐长度统一取值
            cells = row + [""] * (max(mapping.values()) + 1 - len(row)) if row else row
            if not any(cell.strip() for cell in cells):
                continue
            if _is_repeated_header_row(cells):
                continue

            def cell(field: str) -> str:
                column = mapping.get(field)
                if column is None or column >= len(cells):
                    return ""
                return _cell_text(cells[column])

            block_carrier = _is_carrier_block_row(cells, mapping)
            if block_carrier:
                _close_block()
                block_name = block_carrier
                _remember_carrier(
                    carriers,
                    block_carrier,
                    sheet_name,
                    "sheet_name",
                    carrier_column_name,
                    index + 1,
                )
                continue

            carrier_value = cell("carrier")
            destination_value = cell("destination")
            destination_city_value = cell("destination_city")
            origin_province_value = cell("origin_province")
            origin_city_value = cell("origin_city")

            # 线路列写成"起点-终点"时拆开；destination 列也可能是这种写法
            if "route" in mapping and not destination_value and not destination_city_value:
                route_value = cell("route")
                origin_guess, destination_guess = _split_origin_destination(route_value)
                if destination_guess:
                    destination_value = destination_guess
                    if origin_guess and not origin_city_value:
                        origin_city_value = origin_guess
            if destination_value and not destination_city_value and "route" not in mapping:
                origin_guess, destination_guess = _split_origin_destination(destination_value)
                if destination_guess:
                    origin_city_value = origin_city_value or (origin_guess or "")
                    destination_value = destination_guess
                    destination_city_value = destination_guess

            if not destination_value and not destination_city_value:
                continue

            carrier_name = carrier_value or block_name or _sheet_carrier(sheet_name)
            if carrier_name is None:
                carrier_name = sheet_name
                fallback_carrier = True
            else:
                fallback_carrier = False
            if block_name and block_name not in reported_carriers:
                _remember_carrier(
                    carriers,
                    block_name,
                    sheet_name,
                    "sheet_name",
                    carrier_column_name,
                    index + 1,
                )
                reported_carriers.add(block_name)

            if cell("origin"):
                province_guess, city_guess, _ = _split_location(cell("origin"))
                origin_province_value = origin_province_value or province_guess
                origin_city_value = origin_city_value or city_guess

            destination_province_value = cell("destination_province")
            destination_city_value = destination_city_value or destination_value
            province_guess, city_guess, _ = _split_location(destination_city_value)
            destination_province_value = destination_province_value or province_guess
            destination_city_value = city_guess or destination_city_value

            raw_rule_text = ""
            for field in ("continued_price", "first_price", "quote", "price_per_kg"):
                raw_rule_text = f"{raw_rule_text} {cell(field)}"
            raw_rule_text = raw_rule_text.strip()

            fixed_tiers = None
            continued_tiers = None
            for field in ("quote", "continued_price"):
                cell_value = cell(field)
                if not cell_value:
                    continue
                fixed_tiers = fixed_tiers or (
                    _parse_fixed_tiers(cell_value)
                    if any(keyword in cell_value for keyword in FIXED_TIER_KEYWORDS)
                    else None
                )
                continued_tiers = continued_tiers or (
                    _parse_tiers(cell_value)
                    if any(keyword in cell_value for keyword in BANDED_KEYWORDS)
                    or bool(re.search(r"\d\s*(?:-|~|—|–|至|到)\s*\d", cell_value))
                    else None
                )

            # 未映射的列（例如"重量分档""分档价"）也可能是价格说明，逐列兜底尝试
            for column, value in enumerate(cells):
                if column in recognized_columns:
                    continue
                cell_value = _cell_text(value)
                if not cell_value:
                    continue
                if fixed_tiers is None:
                    fixed_tiers = _parse_fixed_tiers(cell_value)
                if continued_tiers is None:
                    continued_tiers = _parse_tiers(cell_value)
                if continued_tiers is not None and fixed_tiers is not None:
                    break

            first_weight_value = _to_float(cell("first_weight_kg"))
            minimum_price_value = None
            for column, value in enumerate(cells):
                if column in recognized_columns:
                    continue
                minimum_price_value = _parse_min_price(value)
                if minimum_price_value is not None:
                    break
            if continued_tiers:
                # 有首重时，分档价通常只作用于续重部分
                default_basis = "continued" if first_weight_value else "total"
                for tier in continued_tiers:
                    if not tier.get("basis"):
                        tier["basis"] = default_basis

            entry: Dict[str, Any] = {
                "id": "",
                "sheet": sheet_name,
                "source_row": index + 1,
                "seller": cell("seller") or None,
                "channel": cell("channel") or None,
                "carrier": carrier_name or None,
                "carrier_source": "carrier_column" if carrier_value else "sheet_name",
                "route": cell("route") or None,
                "eta": cell("eta") or None,
                "origin_province": origin_province_value or None,
                "origin_city": origin_city_value or None,
                "origin": cell("origin") or None,
                "destination_province": destination_province_value or None,
                "destination_city": destination_city_value or None,
                "destination": destination_province_value + destination_city_value or destination_city_value or None,
                "first_weight_kg": first_weight_value,
                "first_price": _to_price(cell("first_price")),
                "continued_unit_kg": _to_float(cell("continued_unit_kg")),
                "continued_price": _to_price(cell("continued_price")),
                "continued_tiers": continued_tiers,
                "fixed_tiers": fixed_tiers,
                "rule_type": None,
                "book_kind": None,
                "quote": _to_price(cell("quote")),
                "min_price": minimum_price_value,
                "confidence": 0.0,
                "review_state": "review",
                "issues": [],
                "raw": {
                    _cell_text(header_cells[column]) or f"col{column}": _cell_text(cells[column])
                    for column in sorted(set(mapping.values()))
                    if column < len(cells)
                } if header_index >= 0 else {},
            }
            column_carrier = ""
            if carrier_value:
                column_carrier = carrier_value
            if column_carrier:
                _remember_carrier(
                    carriers,
                    column_carrier,
                    sheet_name,
                    "carrier_column",
                    carrier_column_name,
                    index + 1,
                )
            entry["book_kind"] = _guess_book_kind(mapping, sheet_name)

            # 只有当"报价"列本身写了重量/续重语义时，才用它推断分档，
            # 纯数字报价则保留为单一报价；无法解析的文本才丢弃
            if entry["quote"] is not None and not (fixed_tiers or continued_tiers):
                quote_text = cell("quote")
                has_semantics = any(
                    keyword in quote_text
                    for keyword in FIXED_TIER_KEYWORDS + BANDED_KEYWORDS + CONTINUED_KEYWORDS
                ) or bool(re.search(r"\d\s*(?:-|~|—|–|至|到|kg|KG|公斤|千克)", quote_text))
                if has_semantics:
                    parsed_fixed = (
                        _parse_fixed_tiers(quote_text)
                        if any(keyword in quote_text for keyword in FIXED_TIER_KEYWORDS)
                        else None
                    )
                    parsed_tiers = entry["continued_tiers"] or _parse_tiers(quote_text)
                    if parsed_fixed or parsed_tiers:
                        entry["quote"] = None
                        entry["fixed_tiers"] = parsed_fixed or entry["fixed_tiers"]
                        entry["continued_tiers"] = parsed_tiers
                    else:
                        # 有计费语义但解析不出具体档位（例如备注文字）→ 不作为稳定价格
                        entry["quote"] = None
                    if any(keyword in quote_text for keyword in FIXED_TIER_KEYWORDS):
                        entry["fixed_tiers"] = _parse_fixed_tiers(quote_text) or entry["fixed_tiers"]
                elif not re.fullmatch(r"[¥￥\s]*\d+(?:\.\d+)?\s*(?:元|块)?", quote_text):
                    # 既没有计费语义、也不是纯金额（例如备注文字）→ 不作为价格
                    entry["quote"] = None

            rule_type, issues, confidence = _infer_rule(entry, raw_rule_text)
            if fallback_carrier:
                issues.append("报价表未写明承运商，已用工作表名占位")
                confidence -= 0.2
            entry["rule_type"] = rule_type
            entry["issues"] = issues
            entry["confidence"] = round(confidence, 2)
            entry["review_state"] = _review_state(issues, confidence)

            row_sequence += 1
            entry["id"] = f"{sheet_name}-{index + 1}-{row_sequence}"
            parsed_rows.append(entry)
            block_rows += 1
            if rule_type:
                block_routes += 1
                block_rule_types.append(rule_type)

        _close_block()
        services.extend(services_found)
        if not services_found:
            warnings.append(f"工作表「{sheet_name}」没有解析到报价行")

    return parsed_rows, list(carriers.values()), services, warnings


def parse_quote_file(data: bytes, filename: str, content_type: Optional[str] = None) -> Dict[str, Any]:
    """把报价表字节内容解析成前端 :class:`LogisticsQuoteParseResponse` 形状。"""
    if not data:
        raise QuoteParseError("上传的文件是空的")
    if len(data) > MAX_UPLOAD_BYTES:
        raise QuoteParseError(
            f"文件过大（{len(data) // 1024} KB），单个报价表请控制在 {MAX_UPLOAD_BYTES // 1024 // 1024} MB 以内"
        )

    extension = ""
    lowered = (filename or "").lower()
    for candidate in SUPPORTED_EXTENSIONS:
        if lowered.endswith(candidate):
            extension = SUPPORTED_EXTENSIONS[candidate]
            break
    if not extension:
        suffix = lowered[lowered.rfind("."):] if "." in lowered else ""
        if suffix in IMAGE_EXTENSIONS:
            raise QuoteParseError("暂不支持图片报价单识别，请上传 Excel 或 CSV 文件")
        if suffix == ".xls":
            raise QuoteParseError("暂不支持旧版 .xls 格式，请在 Excel 中另存为 .xlsx 后重试")
        raise QuoteParseError("只支持 .xlsx / .xlsm / .csv 格式的报价表")

    sheets = _read_tabular(data, extension, filename)
    rows, carriers, services, warnings = _build_parse_rows(sheets)

    valid = sum(1 for row in rows if row["review_state"] == "valid")
    review = sum(1 for row in rows if row["review_state"] == "review")
    rejected = sum(1 for row in rows if row["review_state"] == "rejected")

    book_kind_counts = {"express": 0, "logistics": 0}
    for row in rows:
        if row["book_kind"] in book_kind_counts:
            book_kind_counts[row["book_kind"]] += 1
    book_kind = None
    if any(book_kind_counts.values()):
        book_kind = max(book_kind_counts, key=lambda key: book_kind_counts[key])

    matched_columns: Dict[str, str] = {}
    unmatched_columns: List[str] = []
    for service in services:
        for field, column in service["mapping"].items():
            matched_columns.setdefault(field, column)
    for sheet_name, sheet_rows in sheets:
        if not sheet_rows:
            continue
        for cell in sheet_rows[0][:40]:
            text = _cell_text(cell)
            if not text:
                continue
            if not _match_header_columns([text]):
                if text not in unmatched_columns:
                    unmatched_columns.append(text)
        break

    status = "parsed" if rows and review + rejected == 0 else "needs_review"
    sha256 = hashlib.sha256(data).hexdigest()

    return {
        "success": True,
        "mode": "rate_book_summary" if services else "carrier_only",
        "source": {
            "filename": filename,
            "size": len(data),
            "sha256": sha256,
            "content_type": content_type,
            "file_type": extension,
            "parser_version": PARSER_VERSION,
            "status": status,
        },
        "mapping": {"matched": matched_columns, "unmatched": unmatched_columns[:40]},
        "summary": {"total": len(rows), "valid": valid, "review": review, "rejected": rejected},
        "book_kind": book_kind,
        "service_count": len([service for service in services if service["row_count"]]),
        "route_count": len([row for row in rows if row["rule_type"]]),
        "services": services,
        "carriers": carriers,
        "rows": rows,
        "sample_row": rows[0] if rows else None,
        "warning_count": len(warnings),
        "warnings": warnings[:40],
    }


def create_logistics_quote_router(
    get_current_user: Callable[..., Dict[str, Any]],
    db_manager: Any,
) -> APIRouter:
    router = APIRouter(prefix="/api/logistics", tags=["logistics"])

    def _user_id(current_user: Dict[str, Any]) -> int:
        return int(current_user["user_id"])

    async def _read_upload(file: UploadFile) -> bytes:
        data = await file.read()
        if len(data) > MAX_UPLOAD_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"文件过大，请控制在 {MAX_UPLOAD_BYTES // 1024 // 1024} MB 以内",
            )
        return data

    @router.post("/quote-sources/parse")
    async def parse_quote_source(
        file: UploadFile = File(...),
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        """解析上传的报价表，返回规范化线路与识别摘要（不写入数据库）。"""
        data = await _read_upload(file)
        try:
            result = parse_quote_file(data, file.filename or "quote", file.content_type)
        except QuoteParseError as exc:
            raise HTTPException(
                status_code=400,
                detail={"code": "quote_parse_failed", "message": str(exc)},
            ) from exc
        logger.info(
            f"【quote】user_id={_user_id(current_user)} 解析报价表 {file.filename}："
            f"rows={result['summary']['total']} valid={result['summary']['valid']}"
        )
        return result

    @router.get("/quote-books")
    def list_quote_books(current_user: Dict[str, Any] = Depends(get_current_user)):
        books = db_manager.list_logistics_quote_books(_user_id(current_user))
        return {"success": True, "books": books}

    @router.post("/quote-books")
    async def create_quote_book(
        file: UploadFile = File(...),
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        """解析并保存报价表；同一份文件（sha256 相同）重复上传会覆盖旧结果。"""
        data = await _read_upload(file)
        filename = file.filename or "quote"
        try:
            parsed = parse_quote_file(data, filename, file.content_type)
        except QuoteParseError as exc:
            raise HTTPException(
                status_code=400,
                detail={"code": "quote_parse_failed", "message": str(exc)},
            ) from exc

        user_id = _user_id(current_user)
        source = parsed["source"]
        payload = {key: value for key, value in parsed.items() if key != "success"}

        try:
            saved = db_manager.upsert_logistics_quote_book(
                user_id=user_id,
                filename=source["filename"],
                file_type=source["file_type"],
                size_bytes=source["size"],
                sha256=source["sha256"],
                book_kind=parsed["book_kind"],
                service_count=parsed["service_count"],
                route_count=parsed["route_count"],
                payload=payload,
            )
        except Exception as exc:  # noqa: BLE001 - 入库失败给出可操作提示
            logger.error(f"【quote】保存报价表失败: {exc}")
            raise HTTPException(
                status_code=500,
                detail={"code": "quote_book_save_failed", "message": "报价表保存失败，请稍后重试"},
            ) from exc

        route_import = None
        route_warning = ""
        try:
            route_import = db_manager.sync_logistics_quote_routes(
                user_id=user_id,
                filename=source["filename"],
                file_type=source["file_type"],
                size_bytes=source["size"],
                sha256=source["sha256"],
                book_kind=parsed["book_kind"],
                payload=payload,
                warnings=parsed["warnings"],
            )
        except Exception as exc:  # noqa: BLE001 - 线路同步失败不影响报价表本体
            logger.error(f"【quote】同步报价线路失败: {exc}")
            route_warning = "报价表已保存，但线路明细同步失败，地址匹配功能暂不可用"

        book = db_manager.get_logistics_quote_book(saved["id"], user_id, include_payload=True)
        if not book:
            raise HTTPException(
                status_code=500,
                detail={"code": "quote_book_save_failed", "message": "报价表保存失败，请稍后重试"},
            )
        if isinstance(book.get("payload"), str):
            try:
                book["payload"] = json.loads(book["payload"])
            except (TypeError, ValueError):
                book["payload"] = payload

        if route_import and parsed["route_count"] and not route_import.get("route_count"):
            route_warning = route_warning or "报价表已保存，但没有可参与地址匹配的线路"

        return {
            "success": True,
            "book": book,
            "route_import": (
                {"id": route_import.get("import_id"), "route_count": route_import.get("route_count")}
                if route_import
                else None
            ),
            "route_warning": route_warning,
        }

    @router.get("/quote-books/{book_id}")
    def get_quote_book(
        book_id: int,
        include_payload: bool = True,
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        book = db_manager.get_logistics_quote_book(
            book_id, _user_id(current_user), include_payload=include_payload
        )
        if not book:
            raise HTTPException(status_code=404, detail="报价表不存在")
        if isinstance(book.get("payload"), str):
            try:
                book["payload"] = json.loads(book["payload"])
            except (TypeError, ValueError):
                book["payload"] = {}
        return {"success": True, "book": book}

    @router.delete("/quote-books/{book_id}")
    def delete_quote_book(
        book_id: int,
        current_user: Dict[str, Any] = Depends(get_current_user),
    ):
        if not db_manager.delete_logistics_quote_book(book_id, _user_id(current_user)):
            raise HTTPException(status_code=404, detail="报价表不存在")
        return {"success": True}

    return router


__all__ = ["create_logistics_quote_router", "parse_quote_file", "QuoteParseError", "PARSER_VERSION"]
