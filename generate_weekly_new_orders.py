#!/usr/bin/env python3
"""Generate 「6.周新客订单表」 rows from ERP API + local Excel."""

from __future__ import annotations

import argparse
import csv
import json
import re
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
    file: str = ""


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


def load_traffic_rows(data_dir: Path, a05_path: Path, globs: list[str]) -> list[TrafficRow]:
    rows_out: list[TrafficRow] = []

    def parse_traffic_sheet(rows: list[list[Any]], source_file: str) -> None:
        if len(rows) < 3:
            return
        header = [str(x) for x in rows[1]]
        if "客户姓名" not in "".join(header) and "客户" not in "".join(header):
            return
        # A05 layout: 序号,业务员,客户姓名,注册,国家,添加日期,分级,来源,类型,品类,...
        for row in rows[2:]:
            if len(row) < 6:
                continue
            customer = str(row[2] if len(row) > 2 else "").strip()
            if not customer:
                continue
            add_d = serial_to_date(row[5] if len(row) > 5 else None)
            rows_out.append(
                TrafficRow(
                    sales=str(row[1] if len(row) > 1 else ""),
                    customer=customer,
                    country=str(row[4] if len(row) > 4 else ""),
                    add_date=add_d,
                    source=str(row[7] if len(row) > 7 else ""),
                    traffic_type=str(row[8] if len(row) > 8 else ""),
                    category=str(row[9] if len(row) > 9 else ""),
                    file=source_file,
                )
            )

    a05_sheets = read_xlsx_rows(a05_path)
    for name, rows in a05_sheets.items():
        if "新流量" in name:
            parse_traffic_sheet(rows, f"A05/{name}")

    for pattern in globs:
        for path in sorted(data_dir.glob(pattern)):
            if path.name.startswith("~$"):
                continue
            try:
                for name, rows in read_xlsx_rows(path).items():
                    if "新流量" in name or name.strip() in ("新流量表",):
                        parse_traffic_sheet(rows, path.name)
            except Exception as exc:
                print(f"  warn: skip {path.name}: {exc}", file=sys.stderr)

    return rows_out


def find_traffic(
    traffic_rows: list[TrafficRow],
    sales: str,
    customer: str,
    order_date: date,
) -> TrafficRow | None:
    ns, nc = _norm(sales), _norm(customer)
    candidates = [
        t
        for t in traffic_rows
        if _norm(t.customer) == nc and (_norm(t.sales) == ns or not t.sales)
    ]
    if not candidates:
        candidates = [t for t in traffic_rows if _norm(t.customer) == nc]
    if not candidates:
        return None
    before = [t for t in candidates if t.add_date and t.add_date <= order_date]
    pool = before or candidates
    pool.sort(key=lambda t: t.add_date or date.min, reverse=True)
    # prefer A05
    pool.sort(key=lambda t: (0 if t.file.startswith("A05") else 1))
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

WEEKLY_HEADERS = [
    "周期", "订单日期", "销售人员", "客户姓名", "所属国家", "型号数量",
    "订单总金额（美金）", "实际到帐金额（美金）", "实际到帐金额（人民币）",
    "产品总金额（人民币）", "收款渠道", "采购金额（人民币）", "运费（人民币）",
    "其它运输运费", "物流方式", "初始运费", "毛利润（人民币）", "毛利率",
    "流量添加日期", "流量来源", "咨询品类", "成交产品",
]


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


def write_review_csv(path: Path, headers: list[str], objects: list[Any]) -> None:
    rows = [[getattr(o, h) if hasattr(o, h) else o.get(h) for h in headers] for o in objects]
    write_csv(path, headers, rows)


def patch_weekly_xlsx(xlsx_path: Path, new_rows: list[list[Any]], period: str) -> None:
    """Append/replace rows for target period in sheet 6.周新客订单表."""
    try:
        import openpyxl
    except ImportError:
        print("  skip xlsx writeback: pip install openpyxl", file=sys.stderr)
        return

    wb = openpyxl.load_workbook(xlsx_path)
    sheet_name = next((n for n in wb.sheetnames if "新客订单" in n), None)
    if not sheet_name:
        print("  skip xlsx writeback: sheet not found", file=sys.stderr)
        return
    ws = wb[sheet_name]

    # remove existing rows for same period (col A)
    to_delete = []
    for r in range(4, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and str(val).strip() == period:
            to_delete.append(r)
    for r in reversed(to_delete):
        ws.delete_rows(r, 1)

    start = ws.max_row + 1
    for i, row in enumerate(new_rows):
        for j, val in enumerate(row, 1):
            ws.cell(start + i, j, val)
    wb.save(xlsx_path)
    print(f"  wrote {len(new_rows)} rows to {xlsx_path.name} ({sheet_name})")


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate weekly new-customer order rows")
    parser.add_argument("-c", "--config", default="config.yaml")
    parser.add_argument("--no-xlsx", action="store_true", help="Skip writing back to weekly xlsx")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    cfg = load_config(root / args.config)
    brands = load_brands(root / cfg["paths"]["brands"])
    data_dir = root / cfg["paths"]["data_dir"]
    out_dir = root / cfg["paths"]["output_dir"]

    api = YouroApi(cfg["api"]["base_url"], cfg["api"]["jsessionid"])
    begin = cfg["week"]["begin_date"]
    end = cfg["week"]["end_date"]
    period = cfg["week"].get("period_label") or period_label(parse_api_date(begin), parse_api_date(end))

    print(f"Fetching sales orders {begin} ~ {end} (firstOrder=Y)...")
    sales = api.list_sales_orders(begin, end)
    sales = [s for s in sales if s.get("firstOrder") == "Y" and s.get("company") in (0, "0", None)]
    print(f"  {len(sales)} orders")

    print("Loading Excel sources...")
    a02_match = load_a02_orders(data_dir / cfg["paths"]["a02"])
    traffic_rows = load_traffic_rows(
        data_dir,
        data_dir / cfg["paths"]["a05"],
        cfg["paths"]["salesperson_globs"],
    )
    print(f"  A02 order keys: {len(a02_match)}, traffic rows: {len(traffic_rows)}")

    weekly_rows: list[list[Any]] = []
    purchase_reviews: list[PurchaseReview] = []
    brand_reviews: list[BrandReview] = []

    for sale in sorted(sales, key=lambda s: s.get("orderDate", "")):
        order_no = sale["orderNo"]
        print(f"  processing {order_no}...")
        purchaser = api.get_purchaser_order(order_no)
        od = parse_api_date(sale["orderDate"])
        sales_name = normalize_sales(sale.get("createBy", ""), cfg.get("sales_name_map", {}))
        a02 = a02_match.get((od.isoformat(), _norm(sales_name), _norm(sale.get("customerName", ""))))
        traffic = find_traffic(traffic_rows, sales_name, str(sale.get("customerName", "")), od)

        row, pr, br = build_row(sale, purchaser, a02, traffic, brands, cfg)
        weekly_rows.append(row)
        purchase_reviews.append(pr)
        brand_reviews.append(br)

        if pr.purchase_diff in ("是", "部分缺失", "缺失"):
            print(f"    ⚠ 采购 {order_no}: {pr.purchase_diff} — {pr.purchase_alert}")

    write_csv(out_dir / "6.周新客订单表.csv", WEEKLY_HEADERS, weekly_rows)
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
    print(f"  {out_dir / '6.周新客订单表.csv'}")
    print(f"  {out_dir / '采购核对.csv'}")
    print(f"  {out_dir / '品牌复核.csv'}")

    if not args.no_xlsx:
        xlsx = data_dir / cfg["paths"]["weekly_xlsx"]
        if xlsx.exists():
            patch_weekly_xlsx(xlsx, weekly_rows, period)
        else:
            print(f"  weekly xlsx not found: {xlsx}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
