#!/usr/bin/env python3
"""Generate 「6.周新客订单表」 rows from ERP API + local Excel."""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter
import sys
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config(path: Path) -> dict:
    if yaml is None:
        raise SystemExit("PyYAML required: pip install pyyaml")
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_brands(path: Path) -> dict:
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_exceptions(path: Path) -> dict:
    """单次转化表例外；文件不存在时返回空结构。"""
    empty = {"conversion_excludes": [], "conversion_channel_overrides": []}
    if not path.exists():
        return empty
    with path.open(encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return {
        "conversion_excludes": list(data.get("conversion_excludes") or []),
        "conversion_channel_overrides": list(data.get("conversion_channel_overrides") or []),
    }


def conversion_exclude_order_nos(exceptions: dict) -> set[str]:
    return {str(e["order_no"]) for e in exceptions.get("conversion_excludes", []) if e.get("order_no")}


def conversion_channel_override_map(exceptions: dict) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for entry in exceptions.get("conversion_channel_overrides", []):
        order_no = str(entry.get("order_no") or "").strip()
        channel = str(entry.get("channel") or "").strip().lower()
        if order_no and channel in ("tm", "rfq", "other"):
            out[order_no] = entry
    return out


# ---------------------------------------------------------------------------
# Dates
# ---------------------------------------------------------------------------

def parse_api_date(s: str) -> date:
    return datetime.strptime(s[:10], "%Y-%m-%d").date()


def to_excel_serial(d: date) -> int:
    return (d - date(1899, 12, 30)).days


def serial_to_date(serial: int | float | str) -> date | None:
    try:
        return date(1899, 12, 30) + timedelta(days=int(float(serial)))
    except (TypeError, ValueError):
        return None


def period_label(begin: date, end: date) -> str:
    return f"{begin.month:02d}.{begin.day:02d} - {end.month:02d}.{end.day:02d}"


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

class YouroApi:
    def __init__(self, base_url: str, jsessionid: str):
        self.base_url = base_url.rstrip("/")
        self.cookie = f"JSESSIONID={jsessionid}"

    def _post(self, path: str, data: dict[str, str]) -> dict:
        body = urllib.parse.urlencode(data).encode()
        req = urllib.request.Request(
            f"{self.base_url}{path}",
            data=body,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "X-Requested-With": "XMLHttpRequest",
                "Cookie": self.cookie,
                "User-Agent": "weekly-new-orders/1.0",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            payload = json.loads(resp.read().decode())
        if payload.get("code") != 0:
            raise RuntimeError(f"API error {path}: {payload}")
        return payload

    def list_sales_orders(self, begin: str, end: str, page_size: int = 100) -> list[dict]:
        data = {
            "pageSize": str(page_size),
            "pageNum": "1",
            "orderByColumn": "createTime",
            "isAsc": "desc",
            "orderNo": "",
            "customerId": "",
            "customerName": "",
            "country": "",
            "trackerNo": "",
            "productName": "",
            "shipmentStatus": "",
            "params[beginOrderDate]": begin,
            "params[endOrderDate]": end,
            "createBy": "",
            "purchaser": "",
            "purchaseStatus": "",
            "payMethod": "",
            "currency": "",
            "receiptCheck": "",
            "freightCheck": "",
            "warehouse": "",
            "company": "",
            "isUrgent": "",
            "firstOrder": "Y",
            "statusArrays": "",
        }
        return self._post("/comp/sales-order/list", data).get("rows", [])

    def get_purchaser_order(self, order_no: str) -> dict | None:
        data = {
            "pageSize": "10",
            "pageNum": "1",
            "orderByColumn": "createTime",
            "isAsc": "desc",
            "orderNo": order_no,
            "customerId": "",
            "customerName": "",
            "country": "",
            "trackerNo": "",
            "productName": "",
            "payMethod": "",
            "shipmentStatus": "",
            "params[beginOrderDate]": "2020-01-01",
            "params[endOrderDate]": "2030-12-31",
            "createBy": "",
            "purchaser": "",
            "purchaseStatus": "",
            "warehouse": "",
            "isUrgent": "",
            "firstOrder": "",
            "statusArrays": "",
        }
        rows = self._post("/comp/purchaser-order/list", data).get("rows", [])
        return rows[0] if rows else None


# ---------------------------------------------------------------------------
# Excel (zip/xml — avoids openpyxl compatibility issues on large files)
# ---------------------------------------------------------------------------

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n - 1


def read_xlsx_rows(xlsx_path: Path, sheet_name: str | None = None) -> dict[str, list[list[Any]]]:
    """Return {sheet_name: [row_values,...]} row index 0 = Excel row 1."""
    with zipfile.ZipFile(xlsx_path) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(".//m:si", NS):
                shared.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))

        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = {
            rel.get("Id"): rel.get("Target")
            for rel in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        }
        sheet_paths: dict[str, str] = {}
        for sh in wb.findall(".//m:sheets/m:sheet", NS):
            name = sh.get("name")
            rid = sh.get(REL_NS + "id")
            target = rels.get(rid, "")
            sheet_paths[name] = "xl/" + target.lstrip("/")

        result: dict[str, list[list[Any]]] = {}
        names = [sheet_name] if sheet_name else list(sheet_paths.keys())
        for name in names:
            if name not in sheet_paths:
                continue
            root = ET.fromstring(zf.read(sheet_paths[name]))
            max_col = 0
            sparse: dict[int, dict[int, Any]] = {}
            for row_el in root.findall(".//m:sheetData/m:row", NS):
                r_idx = int(row_el.get("r")) - 1
                sparse[r_idx] = {}
                for c in row_el.findall("m:c", NS):
                    ref = c.get("r", "")
                    m = re.match(r"([A-Z]+)", ref)
                    if not m:
                        continue
                    ci = _col_index(m.group(1))
                    max_col = max(max_col, ci)
                    v_el = c.find("m:v", NS)
                    if v_el is None or v_el.text is None:
                        continue
                    val: Any = v_el.text
                    if c.get("t") == "s":
                        val = shared[int(val)]
                    elif re.fullmatch(r"-?\d+(\.\d+)?", val):
                        val = float(val) if "." in val else int(val)
                    sparse[r_idx][ci] = val

            if not sparse:
                result[name] = []
                continue
            last_row = max(sparse.keys())
            rows_out: list[list[Any]] = []
            for ri in range(last_row + 1):
                row_map = sparse.get(ri, {})
                width = max(max_col + 1, max(row_map.keys(), default=-1) + 1)
                rows_out.append([row_map.get(ci, "") for ci in range(width)])
            result[name] = rows_out
        return result


def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip()).lower()


# 交叉核对客户名归一（拼写差异）
TRAFFIC_CUSTOMER_ALIASES = {
    "airo mafra": "jairo mafra",
}

# 地区分布：国家名合并（与周表习惯一致）
COUNTRY_ALIASES = {
    "印尼": "印度尼西亚",
    "沙特": "沙特阿拉伯",
    "孟加拉": "孟加拉国",
}

# 地区分布 · 印孟巴（印度 + 孟加拉 + 巴基斯坦）
INMENGBA_COUNTRIES = frozenset({"印度", "孟加拉国", "巴基斯坦"})


def _norm_traffic_customer(customer: str) -> str:
    n = _norm(customer)
    return TRAFFIC_CUSTOMER_ALIASES.get(n, n)


def _find_sheet(rows_by_sheet: dict[str, list[list[Any]]], *keywords: str) -> tuple[str, list[list[Any]]]:
    for name, rows in rows_by_sheet.items():
        if all(k in name for k in keywords):
            return name, rows
    for name, rows in rows_by_sheet.items():
        if any(k in name for k in keywords):
            return name, rows
    raise KeyError(f"No sheet matching {keywords}")


@dataclass
class A02OrderRow:
    order_date: date | None = None
    sales: str = ""
    customer: str = ""
    shop: str = ""
    country: str = ""
    qty: int | float | None = None
    order_amount_usd: float | None = None
    payment_usd: float | None = None
    payment_rmb: float | None = None
    product_rmb: float | None = None
    pay_channel: str = ""
    purchase_rmb: float | None = None
    freight_rmb: float | None = None
    other_freight_rmb: float | None = None
    logistics: str = ""
    initial_freight: float | None = None
    gross_profit: float | None = None
    gross_margin: float | None = None


@dataclass
class TrafficRow:
    sales: str = ""
    customer: str = ""
    country: str = ""
    add_date: date | None = None
    source: str = ""
    traffic_type: str = ""
    category: str = ""
    level: str = ""
    shop: str = ""
    file: str = ""
    source_workbook: str = ""
    source_sheet: str = ""


def is_a05_traffic(t: TrafficRow) -> bool:
    if t.source_workbook.startswith("A05"):
        return True
    return t.file.startswith("A05/")


def format_traffic_source(t: TrafficRow) -> str:
    if t.source_workbook and t.source_sheet:
        return f"{t.source_workbook} › {t.source_sheet}"
    if t.source_workbook:
        return t.source_workbook
    return t.file


def summarize_traffic_sources(rows: list[TrafficRow]) -> str:
    if not rows:
        return "（无数据）"
    from collections import Counter
    counts = Counter(format_traffic_source(t) for t in rows)
    return "；".join(f"「{src}」{n}条" for src, n in counts.most_common())


def _float(v: Any) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def load_a02_orders(a02_path: Path) -> dict[tuple, A02OrderRow]:
    """By (date, sales, customer) from 屿路 - 订单信息表."""
    all_sheets = read_xlsx_rows(a02_path)
    _, order_rows = _find_sheet(all_sheets, "订单信息")

    by_match: dict[tuple, A02OrderRow] = {}

    for row in order_rows[3:]:  # skip title/header rows
        if len(row) < 4 or not row[1]:
            continue
        d = serial_to_date(row[1])
        sales = str(row[2] or "")
        customer = str(row[3] or "")
        rec = A02OrderRow(
            order_date=d,
            sales=sales,
            customer=customer,
            shop=str(row[5] if len(row) > 5 else ""),
            country=str(row[6] if len(row) > 6 else ""),
            qty=row[7] if len(row) > 7 else None,
            order_amount_usd=_float(row[8] if len(row) > 8 else None),
            payment_usd=_float(row[10] if len(row) > 10 else None),
            payment_rmb=_float(row[11] if len(row) > 11 else None),
            product_rmb=_float(row[13] if len(row) > 13 else None),
            pay_channel=str(row[14] if len(row) > 14 else ""),
            purchase_rmb=_float(row[15] if len(row) > 15 else None),
            freight_rmb=_float(row[16] if len(row) > 16 else None),
            other_freight_rmb=_float(row[17] if len(row) > 17 else None),
            logistics=str(row[18] if len(row) > 18 else ""),
            initial_freight=_float(row[19] if len(row) > 19 else None),
            gross_profit=_float(row[20] if len(row) > 20 else None),
            gross_margin=_float(row[21] if len(row) > 21 else None),
        )
        if d and sales and customer:
            by_match[(d.isoformat(), _norm(sales), _norm(customer))] = rec

    return by_match


MULTI_SHOP_A060X_PREFIXES = frozenset({"A0604", "A0607"})


def build_a05_shop_index(traffic_rows: list[TrafficRow]) -> dict[tuple[str, str, str], str]:
    """(sales, customer, add_date) -> 屿路/镕川 from A05 authoritative sheets."""
    index: dict[tuple[str, str, str], str] = {}
    for t in traffic_rows:
        if not is_a05_traffic(t) or not t.add_date or not t.shop:
            continue
        key = (_norm(t.sales), _norm(t.customer), t.add_date.isoformat())
        index[key] = t.shop
    return index


def infer_a060x_shop(
    workbook_prefix: str,
    sales: str,
    customer: str,
    add_date: date | None,
    traffic_type: str,
    a05_index: dict[tuple[str, str, str], str],
    default_shop: str,
) -> str:
    """A0604 Grace / A0607 Lily 单表含 Youro+RonChamp，按 A05 对齐 + 类型兜底。"""
    if add_date:
        key = (_norm(sales), _norm(customer), add_date.isoformat())
        if key in a05_index:
            return a05_index[key]

    tt = str(traffic_type or "").strip().upper()
    if workbook_prefix == "A0604" or _norm(sales) == _norm("Grace"):
        if tt == "R-TM" or "R-TM" in tt:
            return "镕川"
        if tt == "TM":
            return "屿路"
    if workbook_prefix == "A0607" or _norm(sales) == _norm("Lily"):
        if tt in ("询盘", "RFQ"):
            return "镕川"
        if tt == "TM":
            return "镕川"
    return default_shop


def load_traffic_rows(data_dir: Path, a05_path: Path, globs: list[str]) -> list[TrafficRow]:
    rows_out: list[TrafficRow] = []

    def parse_traffic_sheet(
        rows: list[list[Any]],
        workbook: str,
        sheet: str,
        default_shop: str = "",
        *,
        a05_index: dict[tuple[str, str, str], str] | None = None,
        workbook_prefix: str = "",
    ) -> None:
        if len(rows) < 3:
            return
        header = [str(x) for x in rows[1]]
        if "客户姓名" not in "".join(header) and "客户" not in "".join(header):
            return
        multi_shop = workbook_prefix in MULTI_SHOP_A060X_PREFIXES
        for row in rows[2:]:
            if len(row) < 6:
                continue
            customer = str(row[2] if len(row) > 2 else "").strip()
            if not customer:
                continue
            add_d = serial_to_date(row[5] if len(row) > 5 else None)
            sales = str(row[1] if len(row) > 1 else "")
            traffic_type = str(row[8] if len(row) > 8 else "")
            if multi_shop and a05_index is not None:
                shop = infer_a060x_shop(
                    workbook_prefix, sales, customer, add_d, traffic_type, a05_index, default_shop
                )
            else:
                shop = default_shop
            rows_out.append(
                TrafficRow(
                    sales=sales,
                    customer=customer,
                    country=str(row[4] if len(row) > 4 else ""),
                    add_date=add_d,
                    source=str(row[7] if len(row) > 7 else ""),
                    traffic_type=traffic_type,
                    category=str(row[9] if len(row) > 9 else ""),
                    level=str(row[6] if len(row) > 6 else ""),
                    shop=shop,
                    file=f"{workbook} › {sheet}",
                    source_workbook=workbook,
                    source_sheet=sheet,
                )
            )

    a05_workbook = a05_path.name
    a05_sheets = read_xlsx_rows(a05_path)
    for name, rows in a05_sheets.items():
        if "新流量" in name:
            shop = "屿路" if "屿路" in name else "镕川" if "镕川" in name else ""
            parse_traffic_sheet(rows, a05_workbook, name, shop)

    a05_index = build_a05_shop_index(rows_out)

    shop_from_file = {
        "A0602": "屿路", "A0603": "屿路", "A0605": "屿路",
        "A0601": "屿路", "A0606": "镕川",
        # A0604 Grace / A0607 Lily：表内混 Youro+RonChamp，用 infer_a060x_shop
        "A0604": "屿路", "A0607": "镕川",
    }
    for pattern in globs:
        for path in sorted(data_dir.glob(pattern)):
            if path.name.startswith("~$"):
                continue
            prefix = path.name.split("-")[0]
            default_shop = shop_from_file.get(prefix, "")
            try:
                for name, rows in read_xlsx_rows(path).items():
                    if "新流量" in name or name.strip() in ("新流量表",):
                        parse_traffic_sheet(
                            rows,
                            path.name,
                            name,
                            default_shop,
                            a05_index=a05_index if prefix in MULTI_SHOP_A060X_PREFIXES else None,
                            workbook_prefix=prefix,
                        )
            except Exception as exc:
                print(f"  warn: skip {path.name}: {exc}", file=sys.stderr)

    return rows_out


def find_traffic(
    traffic_rows: list[TrafficRow],
    sales: str,
    customer: str,
    order_date: date,
    store: str | None = None,
) -> TrafficRow | None:
    ns, nc = _norm(sales), _norm(customer)
    candidates = [
        t
        for t in traffic_rows
        if _norm(t.customer) == nc and (_norm(t.sales) == ns or not t.sales)
    ]
    if store:
        shop = STORE_TO_SHOP.get(store, "")
        store_matched = [t for t in candidates if not t.shop or t.shop == shop or shop_to_store(t.shop) == store]
        if store_matched:
            candidates = store_matched
    if not candidates:
        candidates = [t for t in traffic_rows if _norm(t.customer) == nc]
    if not candidates:
        return None
    before = [t for t in candidates if t.add_date and t.add_date <= order_date]
    pool = before or candidates
    pool.sort(key=lambda t: t.add_date or date.min, reverse=True)
    pool.sort(key=lambda t: (0 if is_a05_traffic(t) else 1))
    return pool[0]


# ---------------------------------------------------------------------------
# Brand resolver
# ---------------------------------------------------------------------------

@dataclass
class BrandResult:
    brand: str
    confidence: str
    source: str


def resolve_brand(product_name: str, category: str, rules: dict) -> BrandResult:
    text = product_name or ""
    upper = text.upper()
    hits: list[tuple[str, str]] = []

    for prefix, brand in (rules.get("prefixes") or {}).items():
        if prefix.upper() in upper:
            hits.append((brand, "prefix"))

    for kw, brand in (rules.get("keywords") or {}).items():
        if kw.upper() in upper or kw in text:
            hits.append((brand, "keyword"))

    cat_brand = None
    for kw, brand in (rules.get("category") or {}).items():
        if kw in (category or ""):
            cat_brand = brand
            break

    hits = [(b, s) for b, s in hits if b != "待确认"]
    if hits:
        brands = {h[0] for h in hits}
        if len(brands) == 1:
            b = next(iter(brands))
            conf = "高" if cat_brand == b else "中"
            return BrandResult(b, conf, hits[0][1])
        return BrandResult("待确认", "低", "conflict")

    if cat_brand:
        return BrandResult(cat_brand, "中", "category")

    return BrandResult("待确认", "待确认", "fallback")


def normalize_sales(name: str, mapping: dict) -> str:
    key = _norm(name)
    if key in mapping:
        return mapping[key]
    return name.strip().title() if name else ""


def pay_channel(api_method: Any, a02: A02OrderRow | None, pay_map: dict) -> str:
    if a02 and a02.pay_channel:
        return a02.pay_channel
    return pay_map.get(str(api_method), "")


# ---------------------------------------------------------------------------
# Row builder
# ---------------------------------------------------------------------------

STORE_YOURO = "Youro"
STORE_RONCHAMP = "RonChamp"
SHOP_TO_STORE = {"屿路": STORE_YOURO, "镕川": STORE_RONCHAMP}
STORE_TO_SHOP = {STORE_YOURO: "屿路", STORE_RONCHAMP: "镕川"}

YOURO_SALES_ORDER = ["Ennerson", "Luck", "Cindy", "Grace", "David", "Lily"]
RONCHAMP_SALES_ORDER = ["Lily", "Sally", "Grace"]


def shop_to_store(shop: str) -> str:
    return SHOP_TO_STORE.get(str(shop or "").strip(), STORE_YOURO)


def store_to_label(store: str) -> str:
    return "RonChamp" if store == STORE_RONCHAMP else "Youro"


def is_l1plus(level: Any) -> bool:
    s = str(level or "").strip().upper().replace(" ", "")
    return s not in ("", "L0") and s != "L1-"


def levels_equivalent(a: str, b: str) -> bool:
    """业务口径：A060x 的 L2 等价于 A05 的 L1+。"""
    na = str(a or "").strip().upper().replace(" ", "")
    nb = str(b or "").strip().upper().replace(" ", "")
    if na == nb:
        return True
    if {na, nb} == {"L1+", "L2"}:
        return True
    return False


def is_alan_assigned(traffic: "TrafficRow") -> bool:
    tt = str(traffic.traffic_type or "").replace(" ", "").lower()
    return "alan分配" in tt


OTHER_CHANNEL_KEYWORDS = ("转介绍", "微信", "公海", "客户介绍", "介绍", "老客户")


def _text_has_other_keyword(text: str) -> bool:
    return any(kw in str(text or "") for kw in OTHER_CHANNEL_KEYWORDS)


def is_traffic_other_channel(traffic: "TrafficRow") -> bool:
    if is_alan_assigned(traffic):
        return True
    combined = f"{traffic.source or ''}{traffic.traffic_type or ''}"
    return _text_has_other_keyword(combined)


def is_sale_other_channel(sale: dict) -> bool:
    for field in ("salesRemark", "remark", "receiptRemark"):
        if _text_has_other_keyword(str(sale.get(field) or "")):
            return True
    return False


def is_explicit_other_channel(traffic: "TrafficRow | None", sale: dict | None = None) -> bool:
    if traffic and is_traffic_other_channel(traffic):
        return True
    if sale and not traffic and is_sale_other_channel(sale):
        return True
    return False


def load_a07_intent_keys(data_dir: Path, a07_filename: str) -> set[tuple[str, str]]:
    """A07 意向订单 → 无 A05 流量时归入「其他」渠道的首单。"""
    path = data_dir / a07_filename
    if not path.exists():
        return set()
    keys: set[tuple[str, str]] = set()
    for sheet, rows in read_xlsx_rows(path).items():
        if "意向" not in sheet:
            continue
        for row in rows[2:]:
            if len(row) < 4:
                continue
            sales = str(row[2] if len(row) > 2 else "").strip()
            customer = str(row[3] if len(row) > 3 else "").strip()
            if sales and customer:
                keys.add((_norm(sales), _norm(customer)))
    return keys


def classify_channel(
    traffic: "TrafficRow | None",
    sale: dict | None = None,
    *,
    sales_name: str = "",
    store: str = "",
    a07_intent: set[tuple[str, str]] | None = None,
) -> str:
    """tm / rfq / other（明确非 TM 来源）/ unclassified（无流量且非明确其他）."""
    if traffic:
        if is_traffic_other_channel(traffic):
            return "other"
        t = str(traffic.traffic_type or "").upper()
        s = str(traffic.source or "").upper()
        if "RFQ" in t or "RFQ" in s:
            return "rfq"
        if t in ("TM", "询盘", "R-TM") or "TM" in t or "询盘" in t:
            return "tm"
        if traffic.source or traffic.traffic_type:
            return "tm"
    if sale and not traffic:
        sn = _norm(sales_name)
        nc = _norm(str(sale.get("customerName", "")))
        if a07_intent and (sn, nc) in a07_intent:
            return "other"
        # 镕川 Sally 公司转介首单：无新流量表记录 → 其他
        if store == STORE_RONCHAMP and sn == _norm("Sally"):
            return "other"
        if is_sale_other_channel(sale):
            return "other"
    return "unclassified"


DUAL_STORE_SALES = {_norm("Grace"), _norm("Lily")}


def resolve_store(
    a02: A02OrderRow | None,
    sale: dict,
    *,
    traffic: TrafficRow | None = None,
    sales_name: str = "",
) -> str:
    """Grace/Lily 双店业务员：有 A05 流量时以流量表店铺为准，否则 A02 所属店铺。"""
    if _norm(sales_name) in DUAL_STORE_SALES and traffic and str(traffic.shop or "").strip():
        return shop_to_store(traffic.shop)
    if a02 and str(a02.shop or "").strip():
        return shop_to_store(a02.shop)
    if traffic and str(traffic.shop or "").strip():
        return shop_to_store(traffic.shop)
    company = sale.get("company")
    if company in (1, "1"):
        return STORE_RONCHAMP
    return STORE_YOURO


@dataclass
class OrderMetrics:
    order_no: str
    store: str
    sales: str
    channel: str
    payment_rmb: float
    gross: float | None
    margin: float | None
    customer: str = ""
    order_date: str = ""
    traffic_file: str = ""
    traffic_source: str = ""


def compute_order_metrics(
    sale: dict,
    purchaser: dict | None,
    a02: A02OrderRow | None,
    traffic: TrafficRow | None,
    brands: dict,
    cfg: dict,
) -> OrderMetrics:
    row, _, _ = build_row(sale, purchaser, a02, traffic, brands, cfg)
    sales = str(row[2])
    payment_rmb = float(row[8]) if row[8] != "" else 0.0
    gross = float(row[16]) if row[16] != "" else None
    margin = float(row[17]) if row[17] != "" else None
    store = resolve_store(a02, sale, traffic=traffic, sales_name=sales)
    channel = classify_channel(
        traffic, sale, sales_name=sales, store=store, a07_intent=cfg.get("_a07_intent")
    )
    return OrderMetrics(
        order_no=str(sale.get("orderNo", "")),
        store=store,
        sales=sales,
        channel=channel,
        payment_rmb=payment_rmb,
        gross=gross,
        margin=margin,
        customer=str(sale.get("customerName", "")),
        order_date=str(sale.get("orderDate", ""))[:10],
        traffic_file=format_traffic_source(traffic) if traffic else "",
        traffic_source=f"{traffic.source or ''}/{traffic.traffic_type or ''}" if traffic else "",
    )


CONVERSION_HEADERS = [
    "所属店铺", "业务员", "总成交数",
    "TM_6月新客数量", "TM_L1+新客数量", "TM_L1+占比",
    "TM_本月成交新客数量", "TM_新客转化率", "TM_新客成交金额", "TM_新客毛利润", "TM_新客毛利率",
    "RFQ_本月成交新客数量", "RFQ_成交金额", "RFQ_新客毛利润", "RFQ_新客毛利率",
    "其他_本月成交新客数量", "其他_成交金额", "其他_新客毛利润", "其他_新客毛利率",
    "备注",
]

WEEKLY_HEADERS = [
    "周期", "订单日期", "销售人员", "客户姓名", "所属国家", "型号数量",
    "订单总金额（美金）", "实际到帐金额（美金）", "实际到帐金额（人民币）",
    "产品总金额（人民币）", "收款渠道", "采购金额（人民币）", "运费（人民币）",
    "其它运输运费", "物流方式", "初始运费", "毛利润（人民币）", "毛利率",
    "流量添加日期", "流量来源", "咨询品类", "成交产品",
]


@dataclass
class ExceptionReview:
    order_no: str
    action: str
    channel: str
    sales: str
    customer: str
    reason: str


@dataclass
class ChannelReview:
    order_no: str
    store: str
    sales: str
    customer: str
    order_date: str
    payment_rmb: float
    gross: float | None
    traffic_file: str
    traffic_source: str
    note: str


@dataclass
class PurchaseReview:
    order_no: str
    purchase_api: float | None
    purchase_a02_order: float | None
    purchase_diff: str
    purchase_alert: str


def compare_purchase(order_no: str, api: float | None, a02: float | None) -> PurchaseReview:
    vals = {"API": api, "A02订单": a02}
    present = {k: v for k, v in vals.items() if v is not None}
    alerts: list[str] = []
    if len(present) < 2:
        diff_flag = "部分缺失" if present else "缺失"
        if len(present) == 1:
            alerts.append(f"仅有一路数据: {list(present.items())[0]}")
    else:
        amounts = list(present.values())
        if max(amounts) - min(amounts) > 0.01:
            diff_flag = "是"
            alerts.append(", ".join(f"{k}({v})" for k, v in present.items()))
        else:
            diff_flag = "否"
    return PurchaseReview(
        order_no=order_no,
        purchase_api=api,
        purchase_a02_order=a02,
        purchase_diff=diff_flag,
        purchase_alert="; ".join(alerts) if alerts else "",
    )


@dataclass
class BrandReview:
    order_no: str
    brand_suggested: str
    brand_confidence: str
    brand_source: str
    productName_raw: str
    category_raw: str
    brand_final: str = ""


def build_row(
    sale: dict,
    purchaser: dict | None,
    a02: A02OrderRow | None,
    traffic: TrafficRow | None,
    brands: dict,
    cfg: dict,
) -> tuple[list[Any], PurchaseReview, BrandReview]:
    order_no = sale["orderNo"]
    order_date = parse_api_date(sale["orderDate"])
    sales = normalize_sales(sale.get("createBy", ""), cfg.get("sales_name_map", {}))
    customer = str(sale.get("customerName", "")).strip()
    country = sale.get("country", "")
    qty = sale.get("qty", "")
    order_amount = float(sale.get("orderAmount") or 0)
    payment_usd = float(sale.get("paymentReceived") or 0)
    product_usd = float(sale.get("productAmount") or 0)
    rate = float(sale.get("exchangeRate") or 6.7)

    payment_rmb = round(payment_usd * rate, 2)
    product_rmb = round(product_usd * rate, 2)

    purchase_api = _float(purchaser.get("purchaseAmount") if purchaser else None)
    purchase_a02 = a02.purchase_rmb if a02 else None
    purchase_review = compare_purchase(order_no, purchase_api, purchase_a02)

    if purchase_api is not None:
        purchase_rmb = purchase_api
    elif purchase_a02 is not None:
        purchase_rmb = purchase_a02
    else:
        purchase_rmb = None

    freight = a02.freight_rmb if a02 else None
    other_freight = a02.other_freight_rmb if a02 else None
    logistics = (a02.logistics if a02 and a02.logistics else "") or str(sale.get("salesRemark") or "").strip()
    # P 列仅写 A02 初始运费；无 A02 时不从 API 回填（与历史周表习惯一致）
    initial_freight = a02.initial_freight if a02 else None

    channel = pay_channel(sale.get("payMethod"), a02, cfg.get("pay_method_map", {}))

    # 毛利润：默认 Q=J−L；仅当 A02 有运费分项时才扩展扣减
    m = freight or 0
    n = other_freight or 0
    o = 0  # 物流方式列为文本，不参与数值扣减
    p = initial_freight or 0
    has_a02_freight = any(v not in (None, "", 0) for v in (freight, other_freight, initial_freight))

    if purchase_rmb is not None:
        if has_a02_freight:
            gross = round(product_rmb - purchase_rmb - m - n - p, 2)
        else:
            gross = round(product_rmb - purchase_rmb, 2)
        margin = round(gross / product_rmb, 4) if product_rmb else None
    else:
        gross = None
        margin = None

    if a02 and a02.gross_profit is not None and gross is not None:
        if abs(a02.gross_profit - gross) / max(abs(gross), 1) < 0.01:
            gross = a02.gross_profit
            margin = a02.gross_margin

    traffic_type = traffic.traffic_type if traffic else ""
    if not traffic_type and traffic:
        traffic_type = traffic.source
    # 无流量匹配时不默认填「转介绍」；仅 classify_channel=other 时在转化逻辑中处理
    category = traffic.category if traffic else ""
    add_serial = to_excel_serial(traffic.add_date) if traffic and traffic.add_date else ""

    product_name = str(sale.get("productName") or "")
    if purchaser and purchaser.get("productName"):
        product_name = product_name or str(purchaser["productName"])
    brand_res = resolve_brand(product_name, category, brands)
    deal_product = f"成交产品：{brand_res.brand}"

    begin = parse_api_date(cfg["week"]["begin_date"])
    end = parse_api_date(cfg["week"]["end_date"])
    period = cfg["week"].get("period_label") or period_label(begin, end)

    row = [
        period,
        to_excel_serial(order_date),
        sales,
        customer,
        country,
        qty,
        order_amount,
        payment_usd,
        payment_rmb,
        product_rmb,
        channel,
        purchase_rmb if purchase_rmb is not None else "",
        freight if freight is not None else "",
        other_freight if other_freight is not None else "",
        logistics,
        initial_freight if initial_freight is not None else "",
        gross if gross is not None else "",
        margin if margin is not None else "",
        add_serial,
        traffic_type,
        category,
        deal_product,
    ]

    brand_review = BrandReview(
        order_no=order_no,
        brand_suggested=brand_res.brand,
        brand_confidence=brand_res.confidence,
        brand_source=brand_res.source,
        productName_raw=product_name[:500],
        category_raw=category,
    )
    return row, purchase_review, brand_review


def write_csv(path: Path, headers: list[str], rows: list[list[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)


def write_conversion_csv(path: Path, title: str, rows: list[list[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([conversion_title_full(title)] + [""] * (len(CONVERSION_HEADERS) - 1))
        w.writerow(CONVERSION_HEADERS)
        w.writerows(rows)


def write_review_csv(path: Path, headers: list[str], objects: list[Any]) -> None:
    rows = [[getattr(o, h) if hasattr(o, h) else o.get(h) for h in headers] for o in objects]
    write_csv(path, headers, rows)


# ---------------------------------------------------------------------------
# Weekly review CSVs (Step 3 traffic / Step 4 business flow)
# ---------------------------------------------------------------------------

def period_slug(begin: date, end: date) -> str:
    return f"{begin.month:02d}{begin.day:02d}-{end.month:02d}{end.day:02d}"


def a07_week_label_candidates(begin: date, end: date, period: str) -> list[str]:
    raw = f"{begin.month}.{begin.day}-{end.month}.{end.day}"
    compact = period.replace(" ", "")
    return list(dict.fromkeys([raw, compact, period]))


INVERTER_CATEGORY_KWS = ("变频器", "inverter", "vfd", "变频", "frequency inverter")

YOURO_BRAND_SHEET = ["台达", "SEW", "西门子", "丹佛斯", "三菱", "Yaskawa", "Omron", "LS", "其它杂类"]
YOURO_BRAND_CAT = {
    "台达": ["台达", "delta"], "SEW": ["sew"], "西门子": ["西门子", "siemens", "电源-si"],
    "丹佛斯": ["丹佛斯", "danfoss"], "三菱": ["三菱", "mitsubishi"],
    "Yaskawa": ["yaskawa", "安川"], "Omron": ["omron", "欧姆龙"], "LS": ["ls", "ls电气"],
}
RONCHAMP_BRAND_SHEET = [
    "CHINT", "Delta", "SEW", "Omron", "SMC", "Festo", "IFM", "Sick", "Baumer",
    "HEIDENHAIN", "Leuze", "Beckhoff", "Konecranes", "ASCO", "其它杂类",
]
RONCHAMP_BRAND_CAT = {
    "CHINT": ["chint", "正泰"], "Delta": ["delta", "台达"], "SEW": ["sew"],
    "Omron": ["omron", "欧姆龙"], "SMC": ["smc"], "Festo": ["festo"], "IFM": ["ifm"],
    "Sick": ["sick"], "Baumer": ["baumer"], "HEIDENHAIN": ["heidenhain"],
    "Leuze": ["leuze"], "Beckhoff": ["beckhoff"], "Konecranes": ["konecranes"], "ASCO": ["asco"],
}

FLOW_ROW_HEADERS = [
    "日期", "TM/询盘新流量", "L1+数量", "L1+占比率", "L3+数量", "L3+占比率",
    "高潜客户数量", "意向订单数量", "新客意向订单数", "新客意向订单金额",
    "新客订单数量", "新客订单金额", "新客毛利润", "备注",
]


def normalize_country(country: str) -> str:
    c = str(country or "").replace("\n", "").strip()
    return COUNTRY_ALIASES.get(c, c)


def aggregate_traffic_countries(traffic_rows: list[TrafficRow]) -> Counter:
    countries: Counter = Counter()
    for t in traffic_rows:
        c = normalize_country(str(t.country or ""))
        if c:
            countries[c] += 1
    return countries


def inmengba_stats(countries: Counter, total: int) -> tuple[int, float | str]:
    cnt = sum(countries.get(c, 0) for c in INMENGBA_COUNTRIES)
    ratio = round(cnt / total, 6) if total else ""
    return cnt, ratio


def is_inverter_category(category: str) -> bool:
    c = str(category or "").lower().replace("\n", " ")
    return any(kw in c for kw in INVERTER_CATEGORY_KWS)


def is_l3plus(level: Any) -> bool:
    s = str(level or "").strip().upper().replace(" ", "")
    return s.startswith("L3") or s.startswith("L4") or s in ("L3+", "L4+")


def is_tm_inquiry_traffic(t: TrafficRow) -> bool:
    if is_alan_assigned(t):
        return False
    tt = str(t.traffic_type or "").upper()
    return "TM" in tt or "询盘" in str(t.traffic_type or "")


def map_traffic_brand(category: str, brand_list: list[str], cat_map: dict[str, list[str]]) -> str:
    c = str(category or "").lower().replace("\n", " ")
    for brand in brand_list:
        if brand == "其它杂类":
            continue
        for kw in cat_map.get(brand, [brand.lower()]):
            if kw in c:
                return brand
    return "其它杂类"


def week_tm_traffic(
    traffic_rows: list[TrafficRow], store: str, week_begin: date, week_end: date
) -> list[TrafficRow]:
    return [
        t
        for t in traffic_rows
        if is_a05_traffic(t)
        and t.add_date
        and week_begin <= t.add_date <= week_end
        and shop_to_store(t.shop or "") == store
        and is_tm_inquiry_traffic(t)
    ]


def parse_a07_week_items(sheet_rows: list[list[Any]], week_labels: list[str]) -> list[dict[str, Any]]:
    """Parse A07 week block; fill-down 业务员 on continuation rows (empty col C)."""
    labels = {lbl.replace(" ", "") for lbl in week_labels}
    items: list[dict[str, Any]] = []
    started = False
    cur_sales = ""
    for row in sheet_rows[3:]:
        d0 = str(row[0] or "").strip()
        d0c = d0.replace(" ", "")
        if d0c in labels:
            started = True
            if len(row) > 2 and str(row[2] or "").strip():
                cur_sales = str(row[2]).strip()
            customer = str(row[3] if len(row) > 3 else "").strip()
            if customer and customer not in ("无", "内容项", "客户名称"):
                items.append({"sales": cur_sales, "row": row})
            continue
        if started and d0 and d0c not in labels:
            if d0.startswith("本周合计"):
                break
            break
        if started:
            if len(row) > 2 and str(row[2] or "").strip():
                cur_sales = str(row[2]).strip()
            customer = str(row[3] if len(row) > 3 else "").strip()
            if customer and customer not in ("无", "内容项", "客户名称"):
                items.append({"sales": cur_sales, "row": row})
    return items


# A07 无店铺列时，业务员默认店（仅流量表未命中时使用）
A07_SALES_DEFAULT_STORE = {
    _norm("Sally"): STORE_RONCHAMP,
    _norm("Ennerson"): STORE_YOURO,
    _norm("Luck"): STORE_YOURO,
    _norm("Cindy"): STORE_YOURO,
    _norm("David"): STORE_YOURO,
}


def infer_a07_row_store(
    sales: str, customer: str, traffic_rows: list[TrafficRow]
) -> tuple[str, str]:
    """Infer Youro/RonChamp from A05 traffic shop; return (store, source)."""
    cn = _norm_traffic_customer(customer)
    sn = _norm(sales)
    stores: list[str] = []
    for t in traffic_rows:
        if not is_a05_traffic(t) or not t.shop:
            continue
        if _norm_traffic_customer(t.customer) != cn:
            continue
        if sn and _norm(t.sales) != sn:
            continue
        stores.append(shop_to_store(t.shop))
    if stores:
        return Counter(stores).most_common(1)[0][0], "traffic"
    cust_stores: list[str] = []
    for t in traffic_rows:
        if not is_a05_traffic(t) or not t.shop:
            continue
        if _norm_traffic_customer(t.customer) == cn:
            cust_stores.append(shop_to_store(t.shop))
    if cust_stores:
        return Counter(cust_stores).most_common(1)[0][0], "traffic_customer"
    if sn in A07_SALES_DEFAULT_STORE:
        return A07_SALES_DEFAULT_STORE[sn], "sales_default"
    return STORE_YOURO, "unmatched"


def _empty_a07_store_stats() -> dict[str, Any]:
    return {
        "intent_count": 0,
        "intent_new_count": 0,
        "intent_new_amount": 0.0,
        "high_potential_count": 0,
        "intent_rows": [],
        "high_potential_rows": [],
        "intent_details": [],
        "high_potential_details": [],
    }


def load_a07_week_stats(
    a07_path: Path, week_labels: list[str], traffic_rows: list[TrafficRow] | None = None
) -> dict[str, Any]:
    empty_store = _empty_a07_store_stats()
    empty: dict[str, Any] = {
        "intent_count": 0,
        "intent_new_count": 0,
        "intent_new_amount": 0.0,
        "high_potential_count": 0,
        "intent_rows": [],
        "high_potential_rows": [],
        "by_store": {
            STORE_YOURO: dict(empty_store),
            STORE_RONCHAMP: dict(empty_store),
        },
        "store_inferences": [],
    }
    if not a07_path.exists():
        return empty
    sheets = read_xlsx_rows(a07_path)
    intent_sheet = next((sheets[k] for k in sheets if "意向" in k and "高潜" not in k), None)
    hp_sheet = next((sheets[k] for k in sheets if "高潜" in k), None)
    if not intent_sheet:
        return empty
    traffic_rows = traffic_rows or []
    intent_items = parse_a07_week_items(intent_sheet, week_labels)
    hp_items = parse_a07_week_items(hp_sheet, week_labels) if hp_sheet else []
    by_store = {STORE_YOURO: _empty_a07_store_stats(), STORE_RONCHAMP: _empty_a07_store_stats()}
    store_inferences: list[dict[str, str]] = []

    for kind, items in (("intent", intent_items), ("high_potential", hp_items)):
        for item in items:
            sales = item["sales"]
            row = item["row"]
            customer = str(row[3] if len(row) > 3 else "").strip()
            store, src = infer_a07_row_store(sales, customer, traffic_rows)
            store_inferences.append({
                "kind": kind, "sales": sales, "customer": customer, "store": store, "source": src,
            })
            bs = by_store[store]
            if kind == "intent":
                bs["intent_rows"].append(row)
                bs["intent_details"].append({"sales": sales, "row": row, "store_source": src})
                bs["intent_count"] += 1
                if str(row[5] if len(row) > 5 else "").strip() == "新":
                    bs["intent_new_count"] += 1
                    bs["intent_new_amount"] = round(
                        bs["intent_new_amount"] + (_float(row[9]) or 0), 2
                    )
            else:
                bs["high_potential_rows"].append(row)
                bs["high_potential_details"].append({"sales": sales, "row": row, "store_source": src})
                bs["high_potential_count"] += 1

    new_rows = [i["row"] for i in intent_items if str(i["row"][5] if len(i["row"]) > 5 else "").strip() == "新"]
    new_amt = sum(_float(r[9]) or 0 for r in new_rows)
    return {
        "intent_count": len(intent_items),
        "intent_new_count": len(new_rows),
        "intent_new_amount": round(new_amt, 2),
        "high_potential_count": len(hp_items),
        "intent_rows": [i["row"] for i in intent_items],
        "high_potential_rows": [i["row"] for i in hp_items],
        "by_store": by_store,
        "store_inferences": store_inferences,
    }


def summarize_weekly_order_rows(rows: list[list[Any]]) -> tuple[int, float, float]:
    """From WEEKLY_HEADERS rows: count, payment_rmb sum, gross sum."""
    pay = gross = 0.0
    for row in rows:
        if len(row) > 8 and row[8] != "":
            pay += float(row[8])
        if len(row) > 16 and row[16] != "":
            gross += float(row[16])
    return len(rows), round(pay, 2), round(gross, 2)


def _write_step3_brand_csv(
    path: Path,
    period: str,
    brand_list: list[str],
    counts: Counter,
    inverter_in_brand: Counter,
    total: int,
    inverter_product_total: int,
    include_inverter_product_row: bool,
) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["周期", period])
        w.writerow(["周总流量合计", total])
        w.writerow(["品牌/产品", "新客流量数量", "其中变频器", "品类流量占比", "备注"])
        for brand in brand_list:
            n = counts.get(brand, 0)
            w.writerow([brand, n, inverter_in_brand.get(brand, 0), round(n / total, 6) if total else "", "按品牌"])
        if include_inverter_product_row:
            w.writerow([
                "变频器", inverter_product_total, "", round(inverter_product_total / total, 6) if total else "",
                "按产品汇总（非品牌）",
            ])


def generate_weekly_review_csvs(
    out_dir: Path,
    period: str,
    week_begin: date,
    week_end: date,
    traffic_rows: list[TrafficRow],
    youro_weekly_rows: list[list[Any]],
    ronchamp_weekly_rows: list[list[Any]],
    a07_path: Path,
) -> list[Path]:
    slug = period_slug(week_begin, week_end)
    a07_labels = a07_week_label_candidates(week_begin, week_end, period)
    a07 = load_a07_week_stats(a07_path, a07_labels, traffic_rows)
    written: list[Path] = []

    for store, prefix, brand_list, cat_map, include_inv_row in (
        (STORE_YOURO, "Youro", YOURO_BRAND_SHEET, YOURO_BRAND_CAT, True),
        (STORE_RONCHAMP, "RonChamp", RONCHAMP_BRAND_SHEET, RONCHAMP_BRAND_CAT, True),
    ):
        tm_rows = week_tm_traffic(traffic_rows, store, week_begin, week_end)
        total = len(tm_rows)
        counts: Counter = Counter()
        inv_in_brand: Counter = Counter()
        for t in tm_rows:
            brand = map_traffic_brand(t.category, brand_list, cat_map)
            counts[brand] += 1
            if is_inverter_category(t.category):
                inv_in_brand[brand] += 1
        inv_product = sum(1 for t in tm_rows if is_inverter_category(t.category))

        brand_path = out_dir / f"Step3-{prefix}-周流量品牌-{slug}.csv"
        _write_step3_brand_csv(
            brand_path, period, brand_list, counts, inv_in_brand, total, inv_product, include_inv_row
        )
        written.append(brand_path)

        other_rows = [t for t in tm_rows if map_traffic_brand(t.category, brand_list, cat_map) == "其它杂类"]
        other_path = out_dir / f"Step3-{prefix}-其它杂类明细-{slug}.csv"
        write_csv(
            other_path,
            ["sales", "customer", "country", "category", "traffic_type", "add_date", "level", "其中变频器"],
            [
                [
                    t.sales, t.customer, str(t.country or "").replace("\n", ""),
                    str(t.category or "").replace("\n", " "), t.traffic_type,
                    t.add_date.isoformat() if t.add_date else "", t.level,
                    "是" if is_inverter_category(t.category) else "",
                ]
                for t in sorted(other_rows, key=lambda x: (x.sales, x.customer))
            ],
        )
        written.append(other_path)

        countries = aggregate_traffic_countries(tm_rows)
        inmengba_cnt, inmengba_ratio = inmengba_stats(countries, total)
        region_path = out_dir / f"Step3-{prefix}-新流量地区-{slug}.csv"
        with region_path.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["周期", period])
            w.writerow(["周总流量合计", total])
            w.writerow(["共几个国家", len(countries)])
            w.writerow(["印孟巴合计", inmengba_cnt, inmengba_ratio])
            w.writerow(["排名", "国家", "数量"])
            for i, (country, cnt) in enumerate(countries.most_common(), 1):
                w.writerow([i, country, cnt])
        written.append(region_path)

    for store, prefix, order_rows in (
        (STORE_YOURO, "Youro", youro_weekly_rows),
        (STORE_RONCHAMP, "RonChamp", ronchamp_weekly_rows),
    ):
        tm_rows = week_tm_traffic(traffic_rows, store, week_begin, week_end)
        tm = len(tm_rows)
        l1 = sum(1 for t in tm_rows if is_l1plus(t.level))
        l3 = sum(1 for t in tm_rows if is_l3plus(t.level))
        order_cnt, order_amt, order_gross = summarize_weekly_order_rows(order_rows)
        bs = a07["by_store"][store]
        flow_path = out_dir / f"Step4-{prefix}-业务流程-{slug}.csv"
        write_csv(
            flow_path,
            FLOW_ROW_HEADERS,
            [[
                period, tm, l1, round(l1 / tm, 6) if tm else "",
                l3, round(l3 / tm, 6) if tm else "",
                bs["high_potential_count"], bs["intent_count"],
                bs["intent_new_count"], bs["intent_new_amount"],
                order_cnt, order_amt, order_gross, "",
            ]],
        )
        written.append(flow_path)

    for store, prefix in ((STORE_YOURO, "Youro"), (STORE_RONCHAMP, "RonChamp")):
        bs = a07["by_store"][store]
        if bs["intent_details"]:
            intent_detail = out_dir / f"Step4-{prefix}-意向订单明细-{slug}.csv"
            write_csv(
                intent_detail,
                ["sales", "customer", "country", "new_or_old", "level", "category", "amount_rmb", "store_source"],
                [
                    [
                        d["sales"],
                        str(r[3] if len(r) > 3 else ""),
                        str(r[4] if len(r) > 4 else ""),
                        str(r[5] if len(r) > 5 else ""),
                        str(r[6] if len(r) > 6 else ""),
                        str(r[8] if len(r) > 8 else ""),
                        r[9] if len(r) > 9 else "",
                        d["store_source"],
                    ]
                    for d in bs["intent_details"]
                    for r in [d["row"]]
                ],
            )
            written.append(intent_detail)
        if bs["high_potential_details"]:
            hp_detail = out_dir / f"Step4-{prefix}-高潜明细-{slug}.csv"
            write_csv(
                hp_detail,
                ["sales", "customer", "country", "level", "category", "amount_rmb", "store_source"],
                [
                    [
                        d["sales"],
                        str(r[3] if len(r) > 3 else ""),
                        str(r[4] if len(r) > 4 else ""),
                        str(r[5] if len(r) > 5 else ""),
                        str(r[7] if len(r) > 7 else ""),
                        r[8] if len(r) > 8 else "",
                        d["store_source"],
                    ]
                    for d in bs["high_potential_details"]
                    for r in [d["row"]]
                ],
            )
            written.append(hp_detail)

    if a07["store_inferences"]:
        infer_path = out_dir / f"Step4-A07-店铺推断-{slug}.csv"
        write_csv(
            infer_path,
            ["kind", "sales", "customer", "store", "source"],
            [
                [inf["kind"], inf["sales"], inf["customer"], inf["store"], inf["source"]]
                for inf in a07["store_inferences"]
            ],
        )
        written.append(infer_path)

    return written


# ---------------------------------------------------------------------------
# Step 6 · Sheet 1 店铺汇总（A03/A04 + 周新客订单）
# ---------------------------------------------------------------------------

SHOP_MONTH_SHEET = "{month}月店铺数据"

STEP6_METRICS = [
    "现有产品数量（个）",
    "在架商品数量（个）",
    "周新上架产品数量（个）",
    "周优化产品数量（个）",
    "现有优品数量（个）",
    "周累计曝光量（个）",
    "周访客人数（个）",
    "周累计点击量（个）",
    "周TM+询盘数量（个）",
    "周L1+数量（个）",
    "周L1+买家占比率",
    "全站推周平均转化成本（元）",
    "全站推周累计转化人数",
    "全站推周曝光",
    "全站推周点击次数",
    "全站推周平均点击成本（元）",
    "周累计自然曝光量（个）",
    "周累计自然点击量（个）",
    "周平均自然点击率",
    "周广告花费（元）",
    "周单个获客成本（元）",
    "截止当前意向订单数",
    "截止当前意向订单金额",
    "周新客订单成交数量（个）",
    "周新客流量成交率",
    "周新客流量成交销售额（元）",
    "周新客流量毛利（元）",
    "周新客流量毛利率",
]


def parse_shop_day_cell(value: Any) -> date | None:
    s = str(value or "").strip()
    if not s or s.startswith("合计"):
        return None
    parts = s.split(".")
    if len(parts) != 2:
        return None
    try:
        return date(2026, int(parts[0]), int(parts[1]))
    except ValueError:
        return None


def _shop_week_rows(
    sheet_rows: list[list[Any]], week_begin: date, week_end: date
) -> list[list[Any]]:
    out: list[list[Any]] = []
    for row in sheet_rows[3:]:
        d = parse_shop_day_cell(row[0] if row else None)
        if d and week_begin <= d <= week_end:
            out.append(row)
    return out


def _shop_sum_col(rows: list[list[Any]], idx: int) -> float:
    return sum(_float(r[idx]) or 0 for r in rows if len(r) > idx)


def _shop_last_col(rows: list[list[Any]], idx: int) -> float | None:
    for row in reversed(rows):
        if len(row) > idx:
            v = _float(row[idx])
            if v is not None:
                return v
    return None


def aggregate_shop_week_stats(
    sheet_rows: list[list[Any]], week_begin: date, week_end: date
) -> dict[str, Any]:
    week = _shop_week_rows(sheet_rows, week_begin, week_end)
    if not week:
        return {}
    tm = _shop_sum_col(week, 13) + _shop_sum_col(week, 14)
    l1 = _shop_sum_col(week, 15)
    l3 = _shop_sum_col(week, 16)
    ad_spend = _shop_sum_col(week, 17)
    qz_spend = _shop_sum_col(week, 27)
    qz_conv = _shop_sum_col(week, 29)
    qz_click = _shop_sum_col(week, 31)
    nat_exp = _shop_sum_col(week, 34)
    nat_click = _shop_sum_col(week, 35)
    return {
        "现有产品数量（个）": _shop_last_col(week, 3),
        "在架商品数量（个）": _shop_last_col(week, 4),
        "周新上架产品数量（个）": _shop_sum_col(week, 5),
        "周优化产品数量（个）": _shop_sum_col(week, 6),
        "现有优品数量（个）": _shop_last_col(week, 7),
        "周累计曝光量（个）": _shop_sum_col(week, 8),
        "周访客人数（个）": _shop_sum_col(week, 11),
        "周累计点击量（个）": _shop_sum_col(week, 9),
        "周TM+询盘数量（个）": tm,
        "周L1+数量（个）": l1,
        "周L1+买家占比率": round(l1 / tm, 12) if tm else "",
        "周L3+数量（个）": l3 if l3 else "",
        "周L3+买家占比率": round(l3 / tm, 12) if tm and l3 else "",
        "全站推周平均转化成本（元）": round(qz_spend / qz_conv, 12) if qz_conv else "",
        "全站推周累计转化人数": qz_conv or "",
        "全站推周曝光": _shop_sum_col(week, 30) or "",
        "全站推周点击次数": qz_click or "",
        "全站推周平均点击成本（元）": round(qz_spend / qz_click, 12) if qz_click else "",
        "周累计自然曝光量（个）": nat_exp or "",
        "周累计自然点击量（个）": nat_click or "",
        "周平均自然点击率": round(nat_click / nat_exp, 12) if nat_exp else "",
        "周广告花费（元）": round(ad_spend, 2) if ad_spend else "",
        "周单个获客成本（元）": round(ad_spend / tm, 12) if tm else "",
    }


def load_shop_week_stats(
    xlsx_path: Path, week_begin: date, week_end: date
) -> dict[str, Any]:
    if not xlsx_path.exists():
        return {}
    sheet_name = SHOP_MONTH_SHEET.format(month=week_end.month)
    sheets = read_xlsx_rows(xlsx_path)
    if sheet_name not in sheets:
        return {}
    return aggregate_shop_week_stats(sheets[sheet_name], week_begin, week_end)


def apply_shop_order_stats(
    stats: dict[str, Any], order_rows: list[list[Any]]
) -> None:
    order_cnt, order_amt, order_gross = summarize_weekly_order_rows(order_rows)
    tm = stats.get("周TM+询盘数量（个）") or 0
    stats["周新客订单成交数量（个）"] = order_cnt
    stats["周新客流量成交率"] = round(order_cnt / tm, 12) if tm else ""
    stats["周新客流量成交销售额（元）"] = order_amt
    stats["周新客流量毛利（元）"] = order_gross
    stats["周新客流量毛利率"] = _margin_pct(order_gross, order_amt) if order_amt else ""


def apply_a07_intent_by_store(
    youro_stats: dict[str, Any],
    ronchamp_stats: dict[str, Any],
    a07_path: Path,
    week_begin: date,
    week_end: date,
    period: str,
    traffic_rows: list[TrafficRow],
) -> None:
    """A07 新客意向按 A05 流量表店铺归属拆分到 Youro / RonChamp。"""
    labels = a07_week_label_candidates(week_begin, week_end, period)
    a07 = load_a07_week_stats(a07_path, labels, traffic_rows)
    ys = a07["by_store"][STORE_YOURO]
    rs = a07["by_store"][STORE_RONCHAMP]
    youro_stats["截止当前意向订单数"] = ys["intent_new_count"] or 0
    youro_stats["截止当前意向订单金额"] = ys["intent_new_amount"] or 0
    ronchamp_stats["截止当前意向订单数"] = rs["intent_new_count"] or 0
    ronchamp_stats["截止当前意向订单金额"] = rs["intent_new_amount"] or 0


STEP7_METRICS = [
    "现有产品数量（个）",
    "在架商品数量（个）",
    "周新上架产品数量（个）",
    "周优化产品数量（个）",
    "现有优品数量（个）",
    "周累计曝光量（个）",
    "周访客人数（个）",
    "周累计点击量（个）",
    "周TM+询盘数量（个）",
    "周L1+数量（个）",
    "周L1+买家占比率",
    "周L3+数量（个）",
    "周L3+买家占比率",
    "周广告花费（元）",
    "周单个获客成本（元）",
]


def build_shop_summary_pair(
    a03_path: Path,
    a04_path: Path,
    a07_path: Path,
    period: str,
    week_begin: date,
    week_end: date,
    youro_order_rows: list[list[Any]],
    ronchamp_order_rows: list[list[Any]],
    traffic_rows: list[TrafficRow],
) -> tuple[dict[str, Any], dict[str, Any]]:
    youro = load_shop_week_stats(a03_path, week_begin, week_end)
    ronchamp = load_shop_week_stats(a04_path, week_begin, week_end)
    apply_shop_order_stats(youro, youro_order_rows)
    apply_shop_order_stats(ronchamp, ronchamp_order_rows)
    apply_a07_intent_by_store(
        youro, ronchamp, a07_path, week_begin, week_end, period, traffic_rows
    )
    return youro, ronchamp


def generate_step6_shop_summary_csv(
    out_dir: Path,
    period: str,
    week_begin: date,
    week_end: date,
    youro: dict[str, Any],
    ronchamp: dict[str, Any],
) -> Path:
    slug = period_slug(week_begin, week_end)
    path = out_dir / f"Step6-店铺汇总-{slug}.csv"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["周期", period, period])
        w.writerow(["指标", "Youro", "RonChamp"])
        for metric in STEP6_METRICS:
            w.writerow([metric, youro.get(metric, ""), ronchamp.get(metric, "")])
    return path


def generate_step7_overview_csv(
    out_dir: Path,
    period: str,
    week_begin: date,
    week_end: date,
    store_label: str,
    stats: dict[str, Any],
) -> Path:
    """Sheet 3（Youro）/ Sheet 1（RonChamp）总览，与 Step6 同店同源。"""
    slug = period_slug(week_begin, week_end)
    path = out_dir / f"Step7-{store_label}-周数据总览-{slug}.csv"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["周期", period])
        w.writerow(["指标", store_label])
        for metric in STEP7_METRICS:
            w.writerow([metric, stats.get(metric, "")])
    return path


def _sales_sort_key(sales: str, order_list: list[str]) -> tuple[int, str]:
    try:
        return (order_list.index(sales), sales)
    except ValueError:
        return (len(order_list), sales)


def _traffic_in_range(t: TrafficRow, store: str, sales: str, month_begin: date, month_end: date) -> bool:
    shop = STORE_TO_SHOP.get(store, "")
    if _norm(t.sales) != _norm(sales):
        return False
    if t.shop and shop_to_store(t.shop) != store and t.shop != shop:
        return False
    if not t.add_date or not (month_begin <= t.add_date <= month_end):
        return False
    return True


def _traffic_row_key(t: TrafficRow) -> tuple[str, str]:
    return (_norm_traffic_customer(t.customer), t.add_date.isoformat() if t.add_date else "")


@dataclass
class TrafficCrossSummary:
    store: str
    sales: str
    period: str
    a05_source: str
    a060x_source: str
    a05_count: int
    a060x_count: int
    overlap: int
    a05_only: int
    a060x_only: int
    field_mismatch: int
    diff: str
    alert: str


@dataclass
class TrafficCrossDetail:
    store: str
    sales: str
    customer: str
    status: str
    add_date_a05: str
    add_date_a060x: str
    level_a05: str
    level_a060x: str
    a05_workbook: str
    a05_sheet: str
    a060x_workbook: str
    a060x_sheet: str
    alert: str


def compare_traffic_cross(
    traffic_rows: list[TrafficRow],
    month_begin: date,
    month_end: date,
    sales_map: dict,
) -> tuple[list[TrafficCrossSummary], list[TrafficCrossDetail]]:
    """Cross-check A05 shop traffic vs A060x salesperson traffic for the same period."""
    period = f"{month_begin.month}.{month_begin.day} - {month_end.month}.{month_end.day}"
    summaries: list[TrafficCrossSummary] = []
    details: list[TrafficCrossDetail] = []

    pairs: set[tuple[str, str]] = set()
    for t in traffic_rows:
        if not t.add_date or not (month_begin <= t.add_date <= month_end):
            continue
        if not t.shop:
            continue
        store = shop_to_store(t.shop)
        sales = normalize_sales(t.sales, sales_map)
        if sales:
            pairs.add((store, sales))

    for store, sales in sorted(pairs):
        a05 = [
            t for t in traffic_rows
            if is_a05_traffic(t) and _traffic_in_range(t, store, sales, month_begin, month_end)
        ]
        a060x = [
            t for t in traffic_rows
            if not is_a05_traffic(t) and _traffic_in_range(t, store, sales, month_begin, month_end)
        ]
        if not a05 and not a060x:
            continue

        a05_source_desc = summarize_traffic_sources(a05)
        a060x_source_desc = summarize_traffic_sources(a060x)

        a05_by_key = {_traffic_row_key(t): t for t in a05}
        a060x_by_key = {_traffic_row_key(t): t for t in a060x}
        overlap_keys = set(a05_by_key) & set(a060x_by_key)
        a05_only_keys = set(a05_by_key) - overlap_keys
        a060x_only_keys = set(a060x_by_key) - overlap_keys

        field_mismatch = 0
        alerts: list[str] = []

        for key in sorted(overlap_keys):
            ta, tb = a05_by_key[key], a060x_by_key[key]
            la, lb = str(ta.level or "").strip(), str(tb.level or "").strip()
            if not levels_equivalent(la, lb):
                field_mismatch += 1
                details.append(
                    TrafficCrossDetail(
                        store=store_to_label(store),
                        sales=sales,
                        customer=ta.customer,
                        status="字段不一致",
                        add_date_a05=ta.add_date.isoformat() if ta.add_date else "",
                        add_date_a060x=tb.add_date.isoformat() if tb.add_date else "",
                        level_a05=la,
                        level_a060x=lb,
                        a05_workbook=ta.source_workbook,
                        a05_sheet=ta.source_sheet,
                        a060x_workbook=tb.source_workbook,
                        a060x_sheet=tb.source_sheet,
                        alert=(
                            f"分级不一致：{format_traffic_source(ta)} 为 {la or '-'}，"
                            f"{format_traffic_source(tb)} 为 {lb or '-'}"
                        ),
                    )
                )

        a05_cust_dates: dict[str, list[TrafficRow]] = {}
        a060x_cust_dates: dict[str, list[TrafficRow]] = {}
        for k in a05_only_keys:
            c = a05_by_key[k].customer
            a05_cust_dates.setdefault(_norm_traffic_customer(c), []).append(a05_by_key[k])
        for k in a060x_only_keys:
            c = a060x_by_key[k].customer
            a060x_cust_dates.setdefault(_norm_traffic_customer(c), []).append(a060x_by_key[k])

        for key in sorted(a05_only_keys):
            ta = a05_by_key[key]
            nc = _norm_traffic_customer(ta.customer)
            if nc in a060x_cust_dates:
                tb = a060x_cust_dates[nc][0]
                details.append(
                    TrafficCrossDetail(
                        store=store_to_label(store),
                        sales=sales,
                        customer=ta.customer,
                        status="日期不一致",
                        add_date_a05=ta.add_date.isoformat() if ta.add_date else "",
                        add_date_a060x=tb.add_date.isoformat() if tb.add_date else "",
                        level_a05=str(ta.level or ""),
                        level_a060x=str(tb.level or ""),
                        a05_workbook=ta.source_workbook,
                        a05_sheet=ta.source_sheet,
                        a060x_workbook=tb.source_workbook,
                        a060x_sheet=tb.source_sheet,
                        alert=(
                            f"同客户两边都有但添加日期不同："
                            f"{format_traffic_source(ta)}={ta.add_date}，"
                            f"{format_traffic_source(tb)}={tb.add_date}"
                        ),
                    )
                )
            else:
                peer = a060x_source_desc if a060x else "（该业务员无 A060x 新流量表）"
                details.append(
                    TrafficCrossDetail(
                        store=store_to_label(store),
                        sales=sales,
                        customer=ta.customer,
                        status="仅A05",
                        add_date_a05=ta.add_date.isoformat() if ta.add_date else "",
                        add_date_a060x="",
                        level_a05=str(ta.level or ""),
                        level_a060x="",
                        a05_workbook=ta.source_workbook,
                        a05_sheet=ta.source_sheet,
                        a060x_workbook="",
                        a060x_sheet="",
                        alert=f"仅存在于 {format_traffic_source(ta)}；对照 {peer} 无此客户",
                    )
                )

        for key in sorted(a060x_only_keys):
            tb = a060x_by_key[key]
            nc = _norm_traffic_customer(tb.customer)
            if nc in a05_cust_dates:
                continue
            peer = a05_source_desc if a05 else "（该店铺无 A05 新流量表）"
            details.append(
                TrafficCrossDetail(
                    store=store_to_label(store),
                    sales=sales,
                    customer=tb.customer,
                    status="仅A060x",
                    add_date_a05="",
                    add_date_a060x=tb.add_date.isoformat() if tb.add_date else "",
                    level_a05="",
                    level_a060x=str(tb.level or ""),
                    a05_workbook="",
                    a05_sheet="",
                    a060x_workbook=tb.source_workbook,
                    a060x_sheet=tb.source_sheet,
                    alert=f"仅存在于 {format_traffic_source(tb)}；对照 {peer} 无此客户",
                )
            )

        a05_n, a060x_n = len(a05), len(a060x)
        if overlap_keys:
            alerts.append(f"客户+日期完全匹配 {len(overlap_keys)} 条")
        if a05_only_keys:
            only_desc = summarize_traffic_sources([a05_by_key[k] for k in a05_only_keys])
            alerts.append(f"仅 A05 侧 {len(a05_only_keys)} 条：{only_desc}")
        if a060x_only_keys:
            only_desc = summarize_traffic_sources([a060x_by_key[k] for k in a060x_only_keys])
            alerts.append(f"仅 A060x 侧 {len(a060x_only_keys)} 条：{only_desc}")
        if field_mismatch:
            alerts.append(f"重叠但分级不一致 {field_mismatch} 条（见明细）")

        if a05_n == a060x_n == len(overlap_keys) and not a05_only_keys and not a060x_only_keys and not field_mismatch:
            diff = "一致"
        elif not a05 or not a060x:
            diff = "部分缺失"
        else:
            diff = "是"

        summaries.append(
            TrafficCrossSummary(
                store=store_to_label(store),
                sales=sales,
                period=period,
                a05_source=a05_source_desc,
                a060x_source=a060x_source_desc,
                a05_count=a05_n,
                a060x_count=a060x_n,
                overlap=len(overlap_keys),
                a05_only=len(a05_only_keys),
                a060x_only=len(a060x_only_keys),
                field_mismatch=field_mismatch,
                diff=diff,
                alert="；".join(alerts) if alerts else "",
            )
        )

    return summaries, details


def _count_traffic(
    traffic_rows: list[TrafficRow],
    store: str,
    sales: str,
    month_begin: date,
    month_end: date,
    *,
    a05_only: bool = False,
) -> tuple[int, int]:
    shop = STORE_TO_SHOP.get(store, "")
    ns = _norm(sales)
    matched = [
        t
        for t in traffic_rows
        if _norm(t.sales) == ns
        and (not t.shop or t.shop == shop or shop_to_store(t.shop) == store)
        and t.add_date
        and month_begin <= t.add_date <= month_end
        and (not a05_only or is_a05_traffic(t))
    ]
    total = len(matched)
    l1plus = sum(1 for t in matched if is_l1plus(t.level))
    return total, l1plus


def _ratio(num: float | int | None, den: float | int | None) -> float | str:
    if not den:
        return ""
    return round(float(num or 0) / float(den), 4)


def _margin_pct(gross: float | None, amount: float) -> float | str:
    if gross is None or not amount:
        return ""
    return round(gross / amount, 4)


def conversion_range(cfg: dict) -> tuple[date, date, str]:
    """当月累计：月初（week.end_date 所在月 1 日）~ week.end_date。"""
    conv = cfg.get("conversion") or {}
    week_end = parse_api_date(cfg["week"]["end_date"])
    month_begin = (
        parse_api_date(conv["month_begin"])
        if conv.get("month_begin")
        else week_end.replace(day=1)
    )
    month_end = (
        parse_api_date(conv["month_end"])
        if conv.get("month_end")
        else week_end
    )
    title = conv.get("title") or (
        f"{month_begin.month}.{month_begin.day} - {month_end.month}.{month_end.day}"
    )
    return month_begin, month_end, title


def conversion_title_full(title: str) -> str:
    return f"Youro & RonChamp 业务员新客转化汇总（{title}）"


def generate_conversion_table(
    month_sales: list[dict],
    a02_match: dict[tuple, A02OrderRow],
    traffic_rows: list[TrafficRow],
    brands: dict,
    cfg: dict,
    api: YouroApi,
) -> tuple[list[list[Any]], list[OrderMetrics], list[ExceptionReview]]:
    conv = cfg.get("conversion") or {}
    month_begin, month_end, _ = conversion_range(cfg)
    sales_map = cfg.get("sales_name_map", {})
    exceptions = cfg.get("_exceptions") or {}
    exclude_orders = conversion_exclude_order_nos(exceptions)
    channel_overrides = conversion_channel_override_map(exceptions)
    applied_exceptions: list[ExceptionReview] = []

    metrics: list[OrderMetrics] = []
    for sale in month_sales:
        if sale.get("firstOrder") != "Y":
            continue
        od = parse_api_date(sale["orderDate"])
        if od < month_begin or od > month_end:
            continue
        order_no = str(sale.get("orderNo", ""))
        sales_name = normalize_sales(sale.get("createBy", ""), sales_map)
        customer = str(sale.get("customerName", ""))
        if order_no in exclude_orders:
            entry = next(
                (e for e in exceptions.get("conversion_excludes", []) if str(e.get("order_no")) == order_no),
                {},
            )
            applied_exceptions.append(
                ExceptionReview(
                    order_no=order_no,
                    action="exclude",
                    channel="",
                    sales=str(entry.get("sales") or sales_name),
                    customer=str(entry.get("customer") or customer),
                    reason=str(entry.get("reason") or "转化表排除"),
                )
            )
            continue
        a02 = a02_match.get((od.isoformat(), _norm(sales_name), _norm(customer)))
        traffic = find_traffic(traffic_rows, sales_name, customer, od, None)
        purchaser = api.get_purchaser_order(sale["orderNo"])
        m = compute_order_metrics(sale, purchaser, a02, traffic, brands, cfg)
        if order_no in channel_overrides:
            ov = channel_overrides[order_no]
            m.channel = str(ov["channel"]).lower()
            applied_exceptions.append(
                ExceptionReview(
                    order_no=order_no,
                    action="channel_override",
                    channel=m.channel,
                    sales=sales_name,
                    customer=customer,
                    reason=str(ov.get("reason") or ""),
                )
            )
        metrics.append(m)

    rows_out: list[list[Any]] = []
    totals = {"I": 0.0, "J": 0.0, "M": 0.0, "N": 0.0, "Q": 0.0, "R": 0.0}
    count_totals = {"D": 0, "E": 0, "G": 0, "L": 0, "P": 0}

    for store, order_list, label in (
        (STORE_YOURO, YOURO_SALES_ORDER, store_to_label(STORE_YOURO)),
        (STORE_RONCHAMP, RONCHAMP_SALES_ORDER, store_to_label(STORE_RONCHAMP)),
    ):
        sales_names: set[str] = set()
        for t in traffic_rows:
            if not is_a05_traffic(t):
                continue
            shop = t.shop or ""
            if shop_to_store(shop) == store and t.add_date and month_begin <= t.add_date <= month_end:
                sales_names.add(normalize_sales(t.sales, sales_map))
        for m in metrics:
            if m.store == store:
                sales_names.add(m.sales)
        sales_names.discard("")

        sorted_sales = sorted(sales_names, key=lambda s: _sales_sort_key(s, order_list))
        first_in_group = True
        for sales in sorted_sales:
            d, e = _count_traffic(traffic_rows, store, sales, month_begin, month_end, a05_only=True)
            tm = [m for m in metrics if m.store == store and m.sales == sales and m.channel == "tm"]
            rfq = [m for m in metrics if m.store == store and m.sales == sales and m.channel == "rfq"]
            other = [m for m in metrics if m.store == store and m.sales == sales and m.channel == "other"]
            all_deals = [m for m in metrics if m.store == store and m.sales == sales]

            tm_amt = sum(m.payment_rmb for m in tm)
            rfq_amt = sum(m.payment_rmb for m in rfq)
            other_amt = sum(m.payment_rmb for m in other)
            tm_gross = sum(m.gross or 0 for m in tm) if tm else None
            rfq_gross = sum(m.gross or 0 for m in rfq) if rfq else None
            other_gross = sum(m.gross or 0 for m in other) if other else None

            g, l_cnt, p = len(tm), len(rfq), len(other)
            c = len(all_deals)

            row = [
                label if first_in_group else "",
                sales,
                c if c else "",
                d if d else "",
                e if e else "",
                _ratio(e, d),
                g if g else "",
                _ratio(g, d),
                round(tm_amt, 2) if tm_amt else "",
                round(tm_gross, 2) if tm_gross is not None and tm else "",
                _margin_pct(tm_gross, tm_amt) if tm and tm_gross is not None else "",
                l_cnt if l_cnt else "",
                round(rfq_amt, 2) if rfq_amt else "",
                round(rfq_gross, 2) if rfq_gross is not None and rfq else "",
                _margin_pct(rfq_gross, rfq_amt) if rfq and rfq_gross is not None else "",
                p if p else "",
                round(other_amt, 2) if other_amt else "",
                round(other_gross, 2) if other_gross is not None and other else "",
                _margin_pct(other_gross, other_amt) if other and other_gross is not None else "",
                "",
            ]
            rows_out.append(row)
            first_in_group = False

            count_totals["D"] += d
            count_totals["E"] += e
            count_totals["G"] += g
            count_totals["L"] += l_cnt
            count_totals["P"] += p
            for key, col_idx in [("I", 8), ("J", 9), ("M", 12), ("N", 13), ("Q", 16), ("R", 17)]:
                val = row[col_idx]
                if val != "":
                    totals[key] += float(val)

    total_row = [
        "总计", "",
        sum(1 for m in metrics),
        count_totals["D"] or "", count_totals["E"] or "", _ratio(count_totals["E"], count_totals["D"]),
        count_totals["G"] or "", _ratio(count_totals["G"], count_totals["D"]),
        round(totals["I"], 2) if totals["I"] else "",
        round(totals["J"], 2) if totals["J"] else "",
        _margin_pct(totals["J"], totals["I"]) if totals["I"] else "",
        count_totals["L"] or "",
        round(totals["M"], 2) if totals["M"] else "",
        round(totals["N"], 2) if totals["N"] else "",
        _margin_pct(totals["N"], totals["M"]) if totals["M"] else "",
        count_totals["P"] or "",
        round(totals["Q"], 2) if totals["Q"] else "",
        round(totals["R"], 2) if totals["R"] else "",
        _margin_pct(totals["R"], totals["Q"]) if totals["Q"] else "",
        "",
    ]
    rows_out.append(total_row)
    unclassified = [m for m in metrics if m.channel == "unclassified"]
    return rows_out, unclassified, applied_exceptions


def process_week_order(
    sale: dict,
    api: YouroApi,
    a02_match: dict[tuple, A02OrderRow],
    traffic_rows: list[TrafficRow],
    brands: dict,
    cfg: dict,
) -> tuple[str, list[Any], PurchaseReview, BrandReview]:
    order_no = sale["orderNo"]
    purchaser = api.get_purchaser_order(order_no)
    od = parse_api_date(sale["orderDate"])
    sales_name = normalize_sales(sale.get("createBy", ""), cfg.get("sales_name_map", {}))
    customer = str(sale.get("customerName", ""))
    a02 = a02_match.get((od.isoformat(), _norm(sales_name), _norm(customer)))
    traffic = find_traffic(traffic_rows, sales_name, customer, od, None)
    store = resolve_store(a02, sale, traffic=traffic, sales_name=sales_name)
    row, pr, br = build_row(sale, purchaser, a02, traffic, brands, cfg)
    return store, row, pr, br


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate weekly new-customer order rows")
    parser.add_argument("-c", "--config", default="config.yaml")
    parser.add_argument("--no-conversion", action="store_true", help="Skip monthly conversion table")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    cfg = load_config(root / args.config)
    brands = load_brands(root / cfg["paths"]["brands"])
    data_dir = root / cfg["paths"]["data_dir"]
    out_dir = root / cfg["paths"]["output_dir"]
    paths = cfg["paths"]
    cfg["_a07_intent"] = load_a07_intent_keys(
        data_dir, paths.get("a07", "A07-意向和高潜订单表.xlsx")
    )
    exc_path = root / paths.get("exceptions", "exceptions.yaml")
    cfg["_exceptions"] = load_exceptions(exc_path)
    exc_n = len(cfg["_exceptions"].get("conversion_excludes", [])) + len(
        cfg["_exceptions"].get("conversion_channel_overrides", [])
    )
    if exc_n:
        print(f"  conversion exceptions: {exc_n} entries ({exc_path.name})")

    api = YouroApi(cfg["api"]["base_url"], cfg["api"]["jsessionid"])
    begin = cfg["week"]["begin_date"]
    end = cfg["week"]["end_date"]
    period = cfg["week"].get("period_label") or period_label(parse_api_date(begin), parse_api_date(end))

    print(f"Fetching sales orders {begin} ~ {end} (firstOrder=Y)...")
    sales = api.list_sales_orders(begin, end)
    sales = [s for s in sales if s.get("firstOrder") == "Y"]
    print(f"  {len(sales)} first-order rows")

    print("Loading Excel sources...")
    a02_match = load_a02_orders(data_dir / paths["a02"])
    traffic_rows = load_traffic_rows(
        data_dir,
        data_dir / paths["a05"],
        paths["salesperson_globs"],
    )
    print(f"  A02 order keys: {len(a02_match)}, traffic rows: {len(traffic_rows)}")

    youro_rows: list[list[Any]] = []
    ronchamp_rows: list[list[Any]] = []
    purchase_reviews: list[PurchaseReview] = []
    brand_reviews: list[BrandReview] = []

    for sale in sorted(sales, key=lambda s: s.get("orderDate", "")):
        order_no = sale["orderNo"]
        print(f"  processing {order_no}...")
        store, row, pr, br = process_week_order(
            sale, api, a02_match, traffic_rows, brands, cfg
        )
        shop_label = STORE_TO_SHOP.get(store, store)
        print(f"    → {shop_label} ({store_to_label(store)})")
        if store == STORE_RONCHAMP:
            ronchamp_rows.append(row)
        else:
            youro_rows.append(row)
        purchase_reviews.append(pr)
        brand_reviews.append(br)

        if pr.purchase_diff in ("是", "部分缺失", "缺失"):
            print(f"    ⚠ 采购 {order_no}: {pr.purchase_diff} — {pr.purchase_alert}")

    write_csv(out_dir / "6.周新客订单表-Youro.csv", WEEKLY_HEADERS, youro_rows)
    write_csv(out_dir / "4.周新客订单表-RonChamp.csv", WEEKLY_HEADERS, ronchamp_rows)
    write_review_csv(
        out_dir / "采购核对.csv",
        ["order_no", "purchase_api", "purchase_a02_order", "purchase_diff", "purchase_alert"],
        purchase_reviews,
    )
    write_review_csv(
        out_dir / "品牌复核.csv",
        ["order_no", "brand_suggested", "brand_confidence", "brand_source", "productName_raw", "category_raw", "brand_final"],
        brand_reviews,
    )

    print(f"\nOutput:")
    print(f"  {out_dir / '6.周新客订单表-Youro.csv'} ({len(youro_rows)} rows)")
    print(f"  {out_dir / '4.周新客订单表-RonChamp.csv'} ({len(ronchamp_rows)} rows)")
    print(f"  {out_dir / '采购核对.csv'}")
    print(f"  {out_dir / '品牌复核.csv'}")

    week_begin = parse_api_date(begin)
    week_end = parse_api_date(end)
    a07_path = data_dir / paths.get("a07", "A07-意向和高潜订单表.xlsx")
    review_paths = generate_weekly_review_csvs(
        out_dir,
        period,
        week_begin,
        week_end,
        traffic_rows,
        youro_rows,
        ronchamp_rows,
        a07_path,
    )
    print("\nStep 3/4 review CSVs:")
    for p in review_paths:
        print(f"  {p}")

    a03_path = data_dir / paths.get("a03", "A03-26年Youro综合运营数据表.xlsx")
    a04_path = data_dir / paths.get("a04", "A04-26年Ronchamp运营数据表.xlsx")
    youro_shop, ronchamp_shop = build_shop_summary_pair(
        a03_path,
        a04_path,
        a07_path,
        period,
        week_begin,
        week_end,
        youro_rows,
        ronchamp_rows,
        traffic_rows,
    )
    step6_path = generate_step6_shop_summary_csv(
        out_dir, period, week_begin, week_end, youro_shop, ronchamp_shop
    )
    step7_youro = generate_step7_overview_csv(
        out_dir, period, week_begin, week_end, "Youro", youro_shop
    )
    step7_ron = generate_step7_overview_csv(
        out_dir, period, week_begin, week_end, "RonChamp", ronchamp_shop
    )
    print(f"\nStep 6 review CSV:")
    print(f"  {step6_path}")
    print(f"Step 7 review CSVs:")
    print(f"  {step7_youro}")
    print(f"  {step7_ron}")
    if not load_shop_week_stats(a03_path, week_begin, week_end):
        print(f"  ⚠ Youro shop stats empty — check {a03_path.name} › {SHOP_MONTH_SHEET.format(month=week_end.month)}")
    if not load_shop_week_stats(a04_path, week_begin, week_end):
        print(f"  ⚠ RonChamp shop stats empty — check {a04_path.name} › {SHOP_MONTH_SHEET.format(month=week_end.month)}")

    if not args.no_conversion:
        month_begin, month_end, title = conversion_range(cfg)
        mb, me = month_begin.isoformat(), month_end.isoformat()
        print(f"\nGenerating conversion table {conversion_title_full(title)}...")
        month_sales = api.list_sales_orders(mb, me)
        conversion_rows, unclassified, applied_exceptions = generate_conversion_table(
            month_sales, a02_match, traffic_rows, brands, cfg, api
        )
        write_conversion_csv(out_dir / "2.新客转化表.csv", title, conversion_rows)
        print(f"  {out_dir / '2.新客转化表.csv'} ({len(conversion_rows)} rows, {mb} ~ {me})")
        if applied_exceptions:
            write_review_csv(
                out_dir / "转化表例外应用.csv",
                ["order_no", "action", "channel", "sales", "customer", "reason"],
                applied_exceptions,
            )
            print(f"  {out_dir / '转化表例外应用.csv'} ({len(applied_exceptions)} applied from exceptions.yaml)")
        write_review_csv(
            out_dir / "渠道未归类.csv",
            [
                "order_no", "store", "sales", "customer", "order_date",
                "payment_rmb", "gross", "traffic_file", "traffic_source", "note",
            ],
            [
                ChannelReview(
                    order_no=m.order_no,
                    store=store_to_label(m.store),
                    sales=m.sales,
                    customer=m.customer,
                    order_date=m.order_date,
                    payment_rmb=m.payment_rmb,
                    gross=m.gross,
                    traffic_file=m.traffic_file,
                    traffic_source=m.traffic_source,
                    note="无新流量匹配，未归入TM/RFQ/其他，请人工确认渠道",
                )
                for m in unclassified
            ],
        )
        if unclassified:
            print(f"  {out_dir / '渠道未归类.csv'} ({len(unclassified)} orders — 计入总成交C，不进其他P列)")

        traffic_summaries, traffic_details = compare_traffic_cross(
            traffic_rows, month_begin, month_end, cfg.get("sales_name_map", {})
        )
        write_review_csv(
            out_dir / "流量交叉核对-汇总.csv",
            [
                "store", "sales", "period", "a05_source", "a060x_source",
                "a05_count", "a060x_count", "overlap",
                "a05_only", "a060x_only", "field_mismatch", "diff", "alert",
            ],
            traffic_summaries,
        )
        write_review_csv(
            out_dir / "流量交叉核对-明细.csv",
            [
                "store", "sales", "customer", "status", "add_date_a05", "add_date_a060x",
                "level_a05", "level_a060x",
                "a05_workbook", "a05_sheet", "a060x_workbook", "a060x_sheet", "alert",
            ],
            traffic_details,
        )
        print(f"  {out_dir / '流量交叉核对-汇总.csv'} ({len(traffic_summaries)} groups)")
        print(f"  {out_dir / '流量交叉核对-明细.csv'} ({len(traffic_details)} discrepancies)")
        for s in traffic_summaries:
            if s.diff not in ("一致",):
                print(f"    ⚠ 流量 {s.store}/{s.sales}: {s.diff} — {s.alert}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
