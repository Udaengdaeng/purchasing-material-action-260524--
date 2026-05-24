# app.py
# Material ERP Dashboard - material-centered + PO/material/date drilldown
# 실행: streamlit run app.py

from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# 0. Page Config / CSS
# =========================================================

st.set_page_config(page_title="Material ERP Dashboard", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 3.5rem !important; padding-bottom: 2rem !important; }
    header[data-testid="stHeader"] {height: 0px;}
    .big-title { font-size: 2rem; font-weight: 800; margin-bottom: 0.35rem; }
    .sub-text { color:#6b7280; font-size:0.95rem; margin-bottom: 1rem; }
    div[data-testid="stMetricValue"] {font-size: 1.55rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 1. Utilities
# =========================================================

def clean(x: Any) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def norm(x: Any) -> str:
    s = clean(x).lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^0-9a-z가-힣#./+\- ]", "", s)
    return s.strip()


def to_num(x: Any, default: float = 0.0) -> float:
    if x is None:
        return default
    try:
        if pd.isna(x):
            return default
    except Exception:
        pass
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("$", "").replace("%", "")
    if s == "" or s.lower() in ["nan", "none"]:
        return default
    try:
        return float(s)
    except Exception:
        return default


def parse_date(x: Any) -> Optional[pd.Timestamp]:
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    dt = pd.to_datetime(x, errors="coerce")
    if pd.isna(dt):
        return None
    return pd.Timestamp(dt).normalize()


def fmt_qty(x: Any) -> str:
    n = to_num(x, 0)
    if abs(n - round(n)) < 1e-9:
        return f"{n:,.0f}"
    return f"{n:,.2f}"


def fmt_money(x: Any) -> str:
    return f"${to_num(x, 0):,.0f}"


def find_sheet(sheet_names: List[str], candidates: List[str]) -> Optional[str]:
    for c in candidates:
        for s in sheet_names:
            if c in s:
                return s
    return None


def parse_leadtime_days(s: Any) -> int:
    """텍스트 형태의 리드타임을 일(day) 단위 정수로 변환. 예: '4weeks' -> 28."""
    if s is None:
        return 0
    try:
        if pd.isna(s):
            return 0
    except Exception:
        pass
    if isinstance(s, (int, float, np.number)):
        return int(s)
    txt = str(s).strip().lower()
    if not txt:
        return 0
    m = re.search(r"(\d+(?:\.\d+)?)\s*(week|wk|w)", txt)
    if m:
        return int(round(float(m.group(1)) * 7))
    m = re.search(r"(\d+(?:\.\d+)?)\s*(month|mon|mo)", txt)
    if m:
        return int(round(float(m.group(1)) * 30))
    m = re.search(r"(\d+(?:\.\d+)?)\s*(day|d)", txt)
    if m:
        return int(round(float(m.group(1))))
    m = re.search(r"(\d+(?:\.\d+)?)", txt)
    if m:
        return int(round(float(m.group(1))))
    return 0


def item_key(x: Any) -> str:
    s = norm(x)
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def status_rank(s: str) -> int:
    return {"부족": 0, "입고대기": 1, "일부입고": 2, "확보됨": 3, "확인필요": 4}.get(s, 9)


def status_label(s: str) -> str:
    return {
        "부족": "🔴 부족",
        "입고대기": "🟣 입고대기",
        "일부입고": "🟠 일부입고",
        "확보됨": "🟢 확보됨",
        "확인필요": "⚪ 확인필요",
    }.get(s, "⚪ 확인필요")


COLOR_MAP = {
    "부족": "#ef4444",
    "입고대기": "#8b5cf6",
    "일부입고": "#f59e0b",
    "확보됨": "#10b981",
    "확인필요": "#9ca3af",
}


# 운송요율은 엑셀의 '운송요율' 시트에서만 가져옵니다.
# 시트가 없거나 데이터가 없으면 운송 리스크 계산은 수행되지 않습니다.
FREIGHT_COLUMNS = ["destination", "mode", "leadtime_days", "rate_usd_per_kg"]
EMPTY_FREIGHT = pd.DataFrame(columns=FREIGHT_COLUMNS)


# =========================================================
# 2. Load Excel
# =========================================================

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> Tuple[List[str], Dict[str, pd.DataFrame]]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets: Dict[str, pd.DataFrame] = {}
    for s in xls.sheet_names:
        try:
            sheets[s] = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=None)
        except Exception:
            sheets[s] = pd.DataFrame()
    return xls.sheet_names, sheets


# =========================================================
# 3. Parsers
# =========================================================

def parse_purchase(raw: pd.DataFrame, tolerance_qty: float = 1.0) -> pd.DataFrame:
    """
    구매데이터 중심 자재 상태.

    핵심 열:
    Y(24) TTL CONS, Z(25) LOSS, AA(26) stock, AC(28) F.O.C,
    AD(29) PO1, AE(30) PO2, AF(31) TT ORDER PO,
    AH(33) BALANCE CONS, AI(34) BALANCE LOSS,
    AP(41) TT RECEIVED, AQ(42) Q'TY SHIPPED, AR(43) ETD, AS(44) ETA

    자재팀 기준:
    - 부족량 = max(Loss 포함 필요량 - (현재 재고 + F.O.C + 발주수량), 0)
    - 미발주와 수량보정필요는 둘 다 '부족'으로 합침
    - 입고 여부는 확보된 자재가 도착했는지 확인하는 보조 상태
    - 출고는 참고값이며 전체 부족 판단에서 단순 차감하지 않음
    """
    cols = [
        "source_row", "supplier", "item", "item_key", "spec", "color", "unit", "unit_price",
        "moq", "leadtime", "ttl_cons", "loss_required", "stock_qty", "foc_qty",
        "po1_qty", "po2_qty", "order_total_qty", "balance_cons_original", "balance_loss_original",
        "balance_cons_calc", "balance_loss_calc", "balance_loss_diff", "received_qty", "shipped_qty",
        "etd", "eta", "remark", "secured_qty", "need_after_stock_qty", "shortage_qty",
        "arrival_gap_qty", "open_order_qty", "coverage_rate", "status", "action", "basis",
        "estimated_shortage_amount", "estimated_total_value", "status_rank"
    ]
    if raw is None or raw.empty or raw.shape[1] < 35:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(12, len(raw)):
        r = raw.iloc[i]
        item = clean(r.iloc[5] if len(r) > 5 else "")
        if not item or norm(item) in ["item", "nan"]:
            continue
        # 더미 행 필터: 자재명이 순수 숫자거나 너무 짧으면 스킵
        if re.fullmatch(r"\d+(\.\d+)?", item) or len(item) < 2:
            continue

        supplier = clean(r.iloc[3] if len(r) > 3 else "")
        if supplier and re.fullmatch(r"\d+(\.\d+)?", supplier):
            continue
        unit = clean(r.iloc[9] if len(r) > 9 else "")
        unit_price = to_num(r.iloc[10] if len(r) > 10 else 0)
        ttl_cons = to_num(r.iloc[24] if len(r) > 24 else 0)
        loss_required = to_num(r.iloc[25] if len(r) > 25 else 0)
        required = loss_required if loss_required > 0 else ttl_cons
        if required <= 0:
            continue

        stock = to_num(r.iloc[26] if len(r) > 26 else 0)
        foc = to_num(r.iloc[28] if len(r) > 28 else 0)
        po1 = to_num(r.iloc[29] if len(r) > 29 else 0)
        po2 = to_num(r.iloc[30] if len(r) > 30 else 0)
        tt_order = to_num(r.iloc[31] if len(r) > 31 else 0)
        order = tt_order if tt_order > 0 else po1 + po2
        bal_cons_orig = to_num(r.iloc[33] if len(r) > 33 else np.nan, np.nan)
        bal_loss_orig = to_num(r.iloc[34] if len(r) > 34 else np.nan, np.nan)
        received = to_num(r.iloc[41] if len(r) > 41 else 0)
        shipped = to_num(r.iloc[42] if len(r) > 42 else 0)

        secured = stock + foc + order
        bal_cons_calc = secured - ttl_cons
        bal_loss_calc = secured - required
        if pd.isna(bal_loss_orig):
            bal_loss_orig = bal_loss_calc
        if pd.isna(bal_cons_orig):
            bal_cons_orig = bal_cons_calc
        bal_loss_diff = bal_loss_calc - bal_loss_orig
        shortage = max(-bal_loss_orig, 0)
        if shortage <= tolerance_qty:
            shortage = 0.0
        need_after_stock = max(required - stock - foc, 0)
        arrival_gap = max(need_after_stock - received, 0)
        open_order = max(order - received, 0)
        coverage = secured / required if required > 0 else 0

        if shortage > 0:
            status = "부족"
            action = f"{supplier or '공급업체 확인'}에 {item} {fmt_qty(shortage)}{unit} 추가 발주 필요"
            basis = "Loss 포함 필요량 > 현재 재고 + F.O.C + 발주수량"
        else:
            if stock + foc >= required - tolerance_qty or stock + foc + received >= required - tolerance_qty:
                status = "확보됨"
                action = "현재 재고/F.O.C/입고량으로 필요량 커버 가능"
                basis = "현재 재고+F.O.C 또는 입고량 기준 확보"
            elif received > tolerance_qty:
                status = "일부입고"
                action = f"{supplier or '공급업체 확인'}에 {item} 잔여 입고 {fmt_qty(max(arrival_gap, 0))}{unit} 확인 필요"
                basis = "발주는 충분하나 일부 입고 상태"
            elif order > tolerance_qty:
                status = "입고대기"
                action = f"{supplier or '공급업체 확인'}에 {item} 발주 {fmt_qty(order)}{unit} 입고/ETA 확인 필요"
                basis = "발주수량은 필요량을 커버하나 입고 미확인"
            else:
                status = "확인필요"
                action = f"{supplier or '공급업체 확인'}에 {item} 구매/입고 상태 확인 필요"
                basis = "필요량은 있으나 재고/발주/입고 정보 확인 필요"

        rows.append({
            "source_row": i + 1,
            "supplier": supplier,
            "item": item,
            "item_key": item_key(item),
            "spec": clean(r.iloc[6] if len(r) > 6 else ""),
            "color": clean(r.iloc[7] if len(r) > 7 else ""),
            "unit": unit,
            "unit_price": unit_price,
            "moq": to_num(r.iloc[13] if len(r) > 13 else 0),
            "leadtime": clean(r.iloc[15] if len(r) > 15 else ""),
            "ttl_cons": ttl_cons,
            "loss_required": required,
            "stock_qty": stock,
            "foc_qty": foc,
            "po1_qty": po1,
            "po2_qty": po2,
            "order_total_qty": order,
            "balance_cons_original": bal_cons_orig,
            "balance_loss_original": bal_loss_orig,
            "balance_cons_calc": bal_cons_calc,
            "balance_loss_calc": bal_loss_calc,
            "balance_loss_diff": bal_loss_diff,
            "received_qty": received,
            "shipped_qty": shipped,
            "etd": parse_date(r.iloc[43] if len(r) > 43 else None),
            "eta": parse_date(r.iloc[44] if len(r) > 44 else None),
            "remark": clean(r.iloc[45] if len(r) > 45 else ""),
            "secured_qty": secured,
            "need_after_stock_qty": need_after_stock,
            "shortage_qty": shortage,
            "arrival_gap_qty": arrival_gap,
            "open_order_qty": open_order,
            "coverage_rate": coverage,
            "status": status,
            "action": action,
            "basis": basis,
            "estimated_shortage_amount": shortage * unit_price,
            "estimated_total_value": required * unit_price,
            "status_rank": status_rank(status),
        })

    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out = out.sort_values(["status_rank", "shortage_qty", "arrival_gap_qty"], ascending=[True, False, False])
    return out


def parse_consumption(raw: pd.DataFrame) -> pd.DataFrame:
    """구매데이터/자재품목리스트의 style별 usage를 PO 상세용으로 long format 변환."""
    cols = [
        "source_row", "style", "style_color", "style_item_code", "supplier", "item", "item_key",
        "spec", "material_color", "unit", "usage_per_unit"
    ]
    if raw is None or raw.empty or raw.shape[0] < 13 or raw.shape[1] < 24:
        return pd.DataFrame(columns=cols)

    style_cols = []
    for c in range(16, min(raw.shape[1], 24)):
        style = clean(raw.iat[1, c] if raw.shape[0] > 1 else "")
        item_code = clean(raw.iat[2, c] if raw.shape[0] > 2 else "")
        color = clean(raw.iat[3, c] if raw.shape[0] > 3 else "")
        if style:
            style_cols.append((c, style, item_code, color))

    rows = []
    for i in range(12, len(raw)):
        r = raw.iloc[i]
        item = clean(r.iloc[5] if len(r) > 5 else "")
        if not item:
            continue
        # 헤더성 더미 행 필터: item/supplier가 순수 숫자거나 너무 짧으면 스킵
        # (구매데이터 시트 마지막 행에 [0,1,2,3,...] 식 더미가 들어있어 자재명='4', 공급업체='2'로 잡히는 케이스 차단)
        supplier_cell = clean(r.iloc[3] if len(r) > 3 else "")
        if re.fullmatch(r"\d+(\.\d+)?", item) or len(item) < 2:
            continue
        if supplier_cell and re.fullmatch(r"\d+(\.\d+)?", supplier_cell):
            continue
        for c, style, item_code, color in style_cols:
            usage = to_num(r.iloc[c] if len(r) > c else 0)
            if usage <= 0:
                continue
            rows.append({
                "source_row": i + 1,
                "style": style,
                "style_color": color,
                "style_item_code": item_code,
                "supplier": clean(r.iloc[3] if len(r) > 3 else ""),
                "item": item,
                "item_key": item_key(item),
                "spec": clean(r.iloc[6] if len(r) > 6 else ""),
                "material_color": clean(r.iloc[7] if len(r) > 7 else ""),
                "unit": clean(r.iloc[9] if len(r) > 9 else ""),
                "usage_per_unit": usage,
            })
    return pd.DataFrame(rows, columns=cols)


def parse_production(raw: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "prod_key", "source_row", "buyer", "prod_year", "prod_month", "ship_month", "pr_no", "buyer_po",
        "item_code", "style", "color", "prod_qty", "sewing_from", "sewing_to", "ship_date",
        "completed_qty", "balance", "completed_pct", "material_status", "remark", "schedule_label"
    ]
    if raw is None or raw.empty or raw.shape[1] < 12:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]
        buyer = clean(r.iloc[1] if len(r) > 1 else "")
        buyer_po = clean(r.iloc[6] if len(r) > 6 else "")
        style = clean(r.iloc[8] if len(r) > 8 else "")
        color = clean(r.iloc[9] if len(r) > 9 else "")
        qty = to_num(r.iloc[10] if len(r) > 10 else 0)
        if not buyer or not style or qty <= 0:
            continue
        if norm(buyer) in ["buyer", "ppm"] or norm(style) == "style":
            continue
        if "total" in norm(style) or "city" in norm(style):
            continue
        prod_key = f"{buyer_po}|{style}|{color}|{qty}|row{i+1}"
        rows.append({
            "prod_key": prod_key,
            "source_row": i + 1,
            "buyer": buyer,
            "prod_year": clean(r.iloc[2] if len(r) > 2 else ""),
            "prod_month": clean(r.iloc[3] if len(r) > 3 else ""),
            "ship_month": clean(r.iloc[4] if len(r) > 4 else ""),
            "pr_no": clean(r.iloc[5] if len(r) > 5 else ""),
            "buyer_po": buyer_po,
            "item_code": clean(r.iloc[7] if len(r) > 7 else ""),
            "style": style,
            "color": color,
            "prod_qty": qty,
            "sewing_from": parse_date(r.iloc[18] if len(r) > 18 else None),
            "sewing_to": parse_date(r.iloc[19] if len(r) > 19 else None),
            "ship_date": parse_date(r.iloc[20] if len(r) > 20 else None),
            "completed_qty": to_num(r.iloc[21] if len(r) > 21 else 0),
            "balance": to_num(r.iloc[22] if len(r) > 22 else 0),
            "completed_pct": to_num(r.iloc[23] if len(r) > 23 else 0),
            "material_status": clean(r.iloc[24] if len(r) > 24 else ""),
            "remark": clean(r.iloc[25] if len(r) > 25 else ""),
            "schedule_label": f"{buyer_po} | {style} / {color}",
        })
    return pd.DataFrame(rows, columns=cols)


def parse_flow(raw: pd.DataFrame, kind: str) -> pd.DataFrame:
    cols = ["kind", "source_row", "buyer", "po", "supplier", "item", "item_key", "color", "unit", "qty", "date"]
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)
    rows = []
    for i in range(5, len(raw)):
        r = raw.iloc[i]
        if kind == "입고":
            buyer = clean(r.iloc[1] if len(r) > 1 else "")
            po = clean(r.iloc[2] if len(r) > 2 else "")
            supplier = clean(r.iloc[4] if len(r) > 4 else "")
            item = clean(r.iloc[6] if len(r) > 6 else "")
            color = clean(r.iloc[8] if len(r) > 8 else "")
            unit = clean(r.iloc[10] if len(r) > 10 else "")
            qty = to_num(r.iloc[12] if len(r) > 12 else 0)
            dt = parse_date(r.iloc[16] if len(r) > 16 else None)
        else:
            buyer = clean(r.iloc[0] if len(r) > 0 else "")
            po = clean(r.iloc[1] if len(r) > 1 else "")
            supplier = clean(r.iloc[3] if len(r) > 3 else "")
            item = clean(r.iloc[5] if len(r) > 5 else "")
            color = clean(r.iloc[7] if len(r) > 7 else "")
            unit = clean(r.iloc[9] if len(r) > 9 else "")
            qty = to_num(r.iloc[11] if len(r) > 11 else 0)
            dt = parse_date(r.iloc[15] if len(r) > 15 else None)
        if not item or qty <= 0:
            continue
        rows.append({
            "kind": kind, "source_row": i + 1, "buyer": buyer, "po": po,
            "supplier": supplier, "item": item, "item_key": item_key(item),
            "color": color, "unit": unit, "qty": qty, "date": dt,
        })
    return pd.DataFrame(rows, columns=cols)


def build_po_materials(selected_prod: pd.Series, consumption: pd.DataFrame, purchase: pd.DataFrame) -> pd.DataFrame:
    if selected_prod is None or consumption.empty:
        return pd.DataFrame()
    style = norm(selected_prod.get("style", ""))
    color = norm(selected_prod.get("color", ""))
    qty = to_num(selected_prod.get("prod_qty", 0))
    matched = consumption[(consumption["style"].apply(norm) == style) & (consumption["style_color"].apply(norm) == color)].copy()
    if matched.empty:
        matched = consumption[consumption["style"].apply(norm) == style].copy()
    if matched.empty:
        return pd.DataFrame()
    matched["po_required_qty"] = matched["usage_per_unit"] * qty
    # 같은 자재명이 색상별로 여러 행 존재할 수 있으므로,
    # 구매데이터와 자재 소요량 정보가 같은 원본 행에서 파생된다는 점을 활용해 source_row 기준으로 우선 매칭한다.
    status_cols = [
        "source_row", "status", "supplier", "loss_required", "stock_qty", "foc_qty", "order_total_qty",
        "secured_qty", "shortage_qty", "received_qty", "arrival_gap_qty", "open_order_qty", "eta", "action", "basis"
    ]
    available_cols = [c for c in status_cols if c in purchase.columns]
    merged = matched.merge(purchase[available_cols], on="source_row", how="left", suffixes=("", "_global"))
    merged["global_status"] = merged["status"].fillna("확인필요")
    merged["global_secured_qty"] = merged["secured_qty"].fillna(0)
    merged["global_shortage_qty"] = merged["shortage_qty"].fillna(0)
    merged["status_label"] = merged["global_status"].apply(status_label)
    return merged


def parse_freight(raw: pd.DataFrame) -> pd.DataFrame:
    """운송요율 시트 파싱. 헤더는 자동 탐지(배송지/destination, mode/운송수단, leadtime/리드타임, rate/단가).
    데이터가 없거나 형식이 맞지 않으면 빈 DataFrame을 반환한다 (운송 리스크 계산 비활성).
    """
    cols = FREIGHT_COLUMNS
    if raw is None or raw.empty:
        return EMPTY_FREIGHT.copy()

    header_row = None
    for i in range(min(10, len(raw))):
        row_txt = " ".join(clean(x).lower() for x in raw.iloc[i].tolist())
        if any(k in row_txt for k in ["destination", "배송지", "도착지"]) and any(
            k in row_txt for k in ["mode", "운송", "type"]
        ):
            header_row = i
            break
    if header_row is None:
        for i in range(min(10, len(raw))):
            row_txt = " ".join(clean(x).lower() for x in raw.iloc[i].tolist())
            if "leadtime" in row_txt or "리드타임" in row_txt:
                header_row = i
                break
    if header_row is None:
        return EMPTY_FREIGHT.copy()

    headers = [clean(x).lower() for x in raw.iloc[header_row].tolist()]
    idx_map: Dict[str, int] = {}
    for j, h in enumerate(headers):
        if not h:
            continue
        if "dest" in h or "배송" in h or "도착" in h:
            idx_map.setdefault("destination", j)
        elif "mode" in h or "운송" in h or "type" in h:
            idx_map.setdefault("mode", j)
        elif "lead" in h or "리드" in h or "day" in h or "일" in h:
            idx_map.setdefault("leadtime_days", j)
        elif "rate" in h or "단가" in h or "price" in h or "$" in h or "usd" in h or "kg" in h:
            idx_map.setdefault("rate_usd_per_kg", j)

    if not all(k in idx_map for k in ["destination", "mode", "leadtime_days", "rate_usd_per_kg"]):
        return EMPTY_FREIGHT.copy()

    rows = []
    for i in range(header_row + 1, len(raw)):
        r = raw.iloc[i]
        dest = clean(r.iloc[idx_map["destination"]] if len(r) > idx_map["destination"] else "")
        mode = clean(r.iloc[idx_map["mode"]] if len(r) > idx_map["mode"] else "").upper()
        lt = to_num(r.iloc[idx_map["leadtime_days"]] if len(r) > idx_map["leadtime_days"] else 0)
        rate = to_num(r.iloc[idx_map["rate_usd_per_kg"]] if len(r) > idx_map["rate_usd_per_kg"] else 0)
        if not dest or not mode or lt <= 0:
            continue
        if mode not in ("SEA", "AIR"):
            if "air" in mode.lower() or "항공" in mode:
                mode = "AIR"
            elif "sea" in mode.lower() or "선박" in mode or "해상" in mode:
                mode = "SEA"
            else:
                continue
        rows.append({"destination": dest.upper(), "mode": mode, "leadtime_days": int(lt), "rate_usd_per_kg": float(rate)})

    if not rows:
        return EMPTY_FREIGHT.copy()
    return pd.DataFrame(rows, columns=cols)


def compute_supply_risk(
    selected_style: str,
    add_qty: float,
    consumption: pd.DataFrame,
    purchase: pd.DataFrame,
    critical_top_n: int = 5,
) -> Dict[str, Any]:
    """공급 리스크 계산.

    - 추가 물량(add_qty) 기준 각 자재의 추가 필요량(usage_per_unit * add_qty)
    - 현재 잉여(secured - loss_required)로 일부 충당, 남는 부분이 부족량
    - 부족량 × 단가 = 추가 주문 비용
    - 리드타임(일) 상위 N개 자재를 주공정으로 식별
    """
    result = {
        "materials": pd.DataFrame(),
        "critical_materials": pd.DataFrame(),
        "total_additional_cost": 0.0,
        "max_leadtime_days": 0,
        "shortage_count": 0,
    }
    if not selected_style or consumption.empty or purchase.empty or add_qty <= 0:
        return result

    style_n = norm(selected_style)
    matched = consumption[consumption["style"].apply(norm) == style_n].copy()
    if matched.empty:
        return result

    # 같은 source_row(자재 행)에서 색상별 usage가 분리되어 있을 수 있어 합산.
    matched["additional_required"] = matched["usage_per_unit"] * add_qty
    agg = matched.groupby("source_row", as_index=False).agg(
        item=("item", "first"),
        supplier=("supplier", "first"),
        unit=("unit", "first"),
        material_color=("material_color", "first"),
        additional_required=("additional_required", "sum"),
    )

    pcols = [
        "source_row", "leadtime", "unit_price", "loss_required", "secured_qty",
        "stock_qty", "foc_qty", "order_total_qty", "received_qty", "status", "eta",
    ]
    pcols = [c for c in pcols if c in purchase.columns]
    merged = agg.merge(purchase[pcols], on="source_row", how="left")
    merged["leadtime_days"] = merged["leadtime"].apply(parse_leadtime_days)
    merged["current_surplus"] = (merged["secured_qty"].fillna(0) - merged["loss_required"].fillna(0)).clip(lower=0)
    merged["new_shortage_qty"] = (merged["additional_required"] - merged["current_surplus"]).clip(lower=0)
    merged["additional_cost"] = merged["new_shortage_qty"] * merged["unit_price"].fillna(0)

    shortage_df = merged[merged["new_shortage_qty"] > 0].copy()
    shortage_df = shortage_df.sort_values("new_shortage_qty", ascending=False)

    # 주공정 자재 = 부족한 자재 중 리드타임이 가장 긴 자재 TOP N
    critical_df = shortage_df.sort_values("leadtime_days", ascending=False).head(critical_top_n).copy()

    result["materials"] = shortage_df.reset_index(drop=True)
    result["critical_materials"] = critical_df.reset_index(drop=True)
    result["total_additional_cost"] = float(shortage_df["additional_cost"].sum())
    result["max_leadtime_days"] = int(shortage_df["leadtime_days"].max()) if not shortage_df.empty else 0
    result["shortage_count"] = int(len(shortage_df))
    return result


def compute_production_risk(
    selected_style: str,
    add_qty: float,
    production: pd.DataFrame,
    daily_capacity: float,
    line_change: bool = False,
    line_change_penalty_days: float = 0.5,
) -> Dict[str, Any]:
    """생산 리스크 계산.

    - 추가 물량 / 일일 생산능력 = 기본 생산일수
    - 라인 체인지 발생 시 +line_change_penalty_days (기본 0.5일) 가산
    - 선택 품목이 기존 생산계획에 있으면 라인 체인지 자동감지 = False (참고용)
    """
    result = {
        "production_leadtime_days": 0.0,
        "base_production_days": 0.0,
        "line_change_penalty_days": 0.0,
        "line_change_required": bool(line_change),
        "line_change_detected": False,
        "existing_runs": pd.DataFrame(),
        "daily_capacity": daily_capacity,
    }
    if add_qty <= 0 or daily_capacity <= 0:
        return result
    base_days = float(add_qty) / float(daily_capacity)
    result["base_production_days"] = round(base_days, 2)

    if not production.empty and selected_style:
        style_n = norm(selected_style)
        existing = production[production["style"].apply(norm) == style_n].copy()
        result["existing_runs"] = existing.reset_index(drop=True)
        result["line_change_detected"] = existing.empty

    penalty = float(line_change_penalty_days) if line_change else 0.0
    result["line_change_penalty_days"] = penalty
    result["production_leadtime_days"] = round(base_days + penalty, 2)
    return result


def compute_transport_risk(
    delivery_date: pd.Timestamp,
    destination: str,
    add_qty: float,
    weight_per_unit_kg: float,
    supply_risk: Dict[str, Any],
    production_risk: Dict[str, Any],
    freight: pd.DataFrame,
) -> Dict[str, Any]:
    """운송 리스크 계산.

    납기까지 남은 일수 = (납기 - 오늘)
    필요 시간 = 자재 조달(주공정 리드타임) + 생산 리드타임 + 운송 리드타임
    SEA로 맞출 수 있으면 SEA 권장, 안되면 AIR 필요.
    """
    result = {
        "options": pd.DataFrame(),
        "air_required": False,
        "recommended_mode": None,
        "days_remaining": 0,
        "material_leadtime_days": int(supply_risk.get("max_leadtime_days", 0)),
        "production_leadtime_days": float(production_risk.get("production_leadtime_days", 0)),
        "total_weight_kg": float(add_qty * weight_per_unit_kg),
        "has_data": False,
    }
    if delivery_date is None or pd.isna(delivery_date):
        return result

    today = pd.Timestamp.today().normalize()
    days_remaining = (pd.Timestamp(delivery_date).normalize() - today).days
    result["days_remaining"] = int(days_remaining)
    pre_transport_days = result["material_leadtime_days"] + result["production_leadtime_days"]

    # 운송요율 데이터가 비어 있으면 계산하지 않고 즉시 반환 (정합성: 임의의 기본값으로 추정 금지)
    if freight is None or freight.empty:
        return result

    dest_n = (destination or "").strip().upper()
    if not dest_n:
        return result
    options = freight[freight["destination"].str.upper() == dest_n].copy()
    if options.empty:
        # 부분 일치만 시도 (없으면 빈 결과 — 임의 fallback 금지)
        options = freight[freight["destination"].str.upper().str.contains(dest_n[:3], na=False)].copy()
    if options.empty:
        return result

    result["has_data"] = True
    options = options.copy()
    options["total_weight_kg"] = result["total_weight_kg"]
    options["transport_cost"] = options["rate_usd_per_kg"] * options["total_weight_kg"]
    options["total_leadtime_days"] = pre_transport_days + options["leadtime_days"]
    options["meets_deadline"] = options["total_leadtime_days"] <= days_remaining
    options["margin_days"] = days_remaining - options["total_leadtime_days"]

    sea_opts = options[options["mode"] == "SEA"]
    air_opts = options[options["mode"] == "AIR"]
    sea_ok = (not sea_opts.empty) and bool(sea_opts["meets_deadline"].any())
    air_ok = (not air_opts.empty) and bool(air_opts["meets_deadline"].any())

    if sea_ok:
        result["recommended_mode"] = "SEA"
        result["air_required"] = False
    elif air_ok:
        result["recommended_mode"] = "AIR"
        result["air_required"] = True
    else:
        result["recommended_mode"] = "AIR (납기 초과)"
        result["air_required"] = True

    result["options"] = options.reset_index(drop=True)
    return result


def build_events(purchase: pd.DataFrame, production: pd.DataFrame, inbound: pd.DataFrame, outbound: pd.DataFrame) -> pd.DataFrame:
    events = []
    if not production.empty:
        for _, r in production.iterrows():
            d = r.get("sewing_from")
            if pd.notna(d):
                events.append({"date": d, "event_type": "생산시작", "item": "", "supplier": "", "qty": 1, "unit": "건", "detail": r.get("schedule_label", "")})
    if not purchase.empty:
        for _, r in purchase.iterrows():
            if pd.notna(r.get("eta")):
                events.append({"date": r.get("eta"), "event_type": "ETA", "item": r.get("item", ""), "supplier": r.get("supplier", ""), "qty": r.get("open_order_qty", 0), "unit": r.get("unit", ""), "detail": r.get("action", "")})
    for df, et in [(inbound, "입고"), (outbound, "출고")]:
        if not df.empty:
            for _, r in df.dropna(subset=["date"]).iterrows():
                events.append({"date": r.get("date"), "event_type": et, "item": r.get("item", ""), "supplier": r.get("supplier", ""), "qty": r.get("qty", 0), "unit": r.get("unit", ""), "detail": r.get("po", "")})
    out = pd.DataFrame(events)
    if out.empty:
        return pd.DataFrame(columns=["date", "event_type", "item", "supplier", "qty", "unit", "detail"])
    out["date"] = pd.to_datetime(out["date"], errors="coerce")
    return out.dropna(subset=["date"])


# =========================================================
# 4. UI
# =========================================================

def show_intro():
    st.markdown('<div class="big-title">📦 Material ERP Dashboard</div>', unsafe_allow_html=True)
    st.info("왼쪽 사이드바에서 엑셀 파일을 업로드하면 분석이 시작됩니다.")
    st.markdown("""
    ### 기준
    - 메인 판단은 자재 기준입니다: `Loss 포함 필요량` vs `현재 재고 + F.O.C + 발주수량`.
    - 미발주/수량보정필요는 자재팀 관점에서 모두 **부족**으로 합산합니다.
    - PO별 화면은 선택한 주문에 필요한 자재의 확보 상태를 확인하는 용도입니다.
    - 날짜별 화면은 전체 이벤트 건수와 날짜/자재별 상세를 함께 보여줍니다.
    """)


def show_dashboard(purchase: pd.DataFrame, production: pd.DataFrame):
    st.markdown('<div class="big-title">📦 Material ERP Dashboard</div>', unsafe_allow_html=True)
    st.caption("자재 부족량 중심으로 전체 현황을 먼저 확인합니다.")
    if purchase.empty:
        st.warning("구매데이터를 읽지 못했습니다.")
        return
    shortage = purchase[purchase["status"] == "부족"]
    waiting = purchase[purchase["status"] == "입고대기"]
    partial = purchase[purchase["status"] == "일부입고"]
    secured = purchase[purchase["status"] == "확보됨"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔴 부족 자재", f"{len(shortage):,}")
    c2.metric("부족 수량 합계", f"{shortage['shortage_qty'].sum():,.1f}")
    c3.metric("🟣 입고대기", f"{len(waiting):,}")
    c4.metric("🟢 확보됨", f"{len(secured):,}")

    left, right = st.columns([0.9, 1.1])
    with left:
        st.subheader("자재 확보 상태")
        cnt = purchase["status"].value_counts().reset_index()
        cnt.columns = ["status", "count"]
        fig = px.pie(cnt, names="status", values="count", hole=0.45, color="status", color_discrete_map=COLOR_MAP)
        fig.update_traces(textposition="inside", textinfo="label+percent")
        fig.update_layout(height=340, margin=dict(l=10, r=10, t=20, b=10), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.subheader("부족 자재 TOP")
        if shortage.empty:
            st.success("부족 자재가 없습니다.")
        else:
            top = shortage.sort_values("shortage_qty", ascending=False).head(12)
            fig = px.bar(top, x="shortage_qty", y="item", orientation="h", text="shortage_qty", hover_data=["supplier", "unit", "loss_required", "secured_qty"])
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("공급업체별 조치 필요")
        contact = purchase[purchase["status"].isin(["부족", "입고대기", "일부입고"])].copy()
        contact["supplier"] = contact["supplier"].replace("", "공급업체 확인")
        if contact.empty:
            st.success("공급업체 확인 항목이 없습니다.")
        else:
            sup = contact.groupby("supplier", as_index=False).agg(
                부족=("status", lambda x: int((x == "부족").sum())),
                입고대기=("status", lambda x: int((x == "입고대기").sum())),
                일부입고=("status", lambda x: int((x == "일부입고").sum())),
                부족수량=("shortage_qty", "sum"),
            )
            sup["총건수"] = sup["부족"] + sup["입고대기"] + sup["일부입고"]
            sup = sup.sort_values(["총건수", "부족수량"], ascending=[False, False]).head(12)
            melted = sup.melt(id_vars="supplier", value_vars=["부족", "입고대기", "일부입고"], var_name="구분", value_name="건수")
            fig = px.bar(melted, x="supplier", y="건수", color="구분", text="건수", color_discrete_map={"부족":"#ef4444", "입고대기":"#8b5cf6", "일부입고":"#f59e0b"})
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10), xaxis_tickangle=-30, legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.subheader("생산 일정 요약")
        if production.empty:
            st.info("생산계획 데이터 없음")
        else:
            d = production.copy()
            d["month"] = pd.to_datetime(d["sewing_from"], errors="coerce").dt.to_period("M").astype(str)
            m = d.groupby("month", as_index=False).agg(생산건수=("buyer_po", "count"), 생산수량=("prod_qty", "sum"))
            fig = px.bar(m, x="month", y="생산건수", text="생산건수")
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)


def show_material_view(purchase: pd.DataFrame):
    st.header("자재별 보기")
    if purchase.empty:
        st.warning("조회 가능한 구매데이터가 없습니다.")
        return
    statuses = ["부족", "입고대기", "일부입고", "확보됨", "확인필요"]
    suppliers = sorted(purchase["supplier"].replace("", "공급업체 확인").dropna().unique().tolist())
    f1, f2, f3 = st.columns([1, 1, 1.8])
    with f1:
        status_filter = st.multiselect("상태", statuses, default=["부족", "입고대기", "일부입고"])
    with f2:
        supplier_filter = st.multiselect("Supplier", suppliers)
    with f3:
        keyword = st.text_input("자재명 / Supplier / Color 검색")
    data = purchase.copy()
    if status_filter:
        data = data[data["status"].isin(status_filter)]
    if supplier_filter:
        data = data[data["supplier"].replace("", "공급업체 확인").isin(supplier_filter)]
    if keyword:
        k = norm(keyword)
        data = data[data["item"].apply(norm).str.contains(k, na=False) | data["supplier"].apply(norm).str.contains(k, na=False) | data["color"].apply(norm).str.contains(k, na=False)]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("자재", f"{len(data):,}")
    c2.metric("부족량", f"{data['shortage_qty'].sum():,.1f}")
    c3.metric("입고대기량", f"{data['open_order_qty'].sum():,.1f}")
    c4.metric("부족 추정금액", fmt_money(data["estimated_shortage_amount"].sum()))
    view = data.copy()
    view["상태"] = view["status"].apply(status_label)
    view["ETA"] = pd.to_datetime(view["eta"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    st.dataframe(view[["상태", "supplier", "item", "spec", "color", "unit", "loss_required", "stock_qty", "foc_qty", "order_total_qty", "secured_qty", "balance_loss_original", "shortage_qty", "received_qty", "open_order_qty", "arrival_gap_qty", "ETA", "action", "basis"]].rename(columns={
        "supplier":"Supplier", "item":"자재명", "spec":"Spec", "color":"Color", "unit":"단위", "loss_required":"Loss 포함 필요량", "stock_qty":"현재 재고", "foc_qty":"F.O.C", "order_total_qty":"발주 수량", "secured_qty":"확보량", "balance_loss_original":"BALANCE LOSS", "shortage_qty":"부족량", "received_qty":"입고 수량", "open_order_qty":"미입고 발주잔량", "arrival_gap_qty":"잔여 입고 확인량", "action":"요청 액션", "basis":"판단 기준"
    }), use_container_width=True, hide_index=True)
    st.download_button("현재 필터 결과 CSV 다운로드", data=data.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"), file_name="material_view_filtered.csv", mime="text/csv")


def show_po_view(production: pd.DataFrame, consumption: pd.DataFrame, purchase: pd.DataFrame):
    st.header("PO별 보기")
    st.caption("주문을 선택하면 해당 주문에 필요한 자재와 전체 확보 상태를 함께 보여줍니다. 자재 소요량 정보가 없는 스타일은 별도로 표시됩니다.")
    if production.empty:
        st.warning("생산계획 데이터가 없습니다.")
        return
    prod = production.copy()
    prod["label"] = prod["schedule_label"] + " | " + prod["prod_qty"].apply(lambda x: f"{x:,.0f}PCS")
    selected = st.selectbox("주문 선택", prod["label"].tolist())
    row = prod.loc[prod["label"] == selected].iloc[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Buyer", row["buyer"])
    c2.metric("생산수량", f"{row['prod_qty']:,.0f} PCS")
    c3.metric("MATERIAL", row["material_status"] if row["material_status"] else "-")
    c4.metric("Ship", pd.to_datetime(row["ship_date"]).strftime("%Y-%m-%d") if pd.notna(row["ship_date"]) else "-")
    detail = build_po_materials(row, consumption, purchase)
    if detail.empty:
        st.warning("이 주문은 현재 자재 소요량 정보가 없어 자재별 확보 여부를 계산할 수 없습니다.")
        return
    d1, d2, d3 = st.columns(3)
    d1.metric("필요 자재", f"{len(detail):,}")
    d2.metric("부족 자재", f"{(detail['global_status'] == '부족').sum():,}")
    d3.metric("입고대기/일부입고", f"{detail['global_status'].isin(['입고대기','일부입고']).sum():,}")
    left, right = st.columns([0.8, 1.2])
    with left:
        cnt = detail["global_status"].value_counts().reset_index()
        cnt.columns = ["status", "count"]
        fig = px.pie(cnt, names="status", values="count", hole=0.45, color="status", color_discrete_map=COLOR_MAP)
        fig.update_layout(height=320, margin=dict(l=10, r=10, t=20, b=10), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        short = detail[detail["global_status"] == "부족"].copy()
        if short.empty:
            st.success("이 주문에 연결된 자재 중 전체 기준 부족 자재가 없습니다.")
        else:
            fig = px.bar(short.sort_values("global_shortage_qty", ascending=False).head(12), x="global_shortage_qty", y="item", orientation="h", text="global_shortage_qty")
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder":"total ascending"})
            st.plotly_chart(fig, use_container_width=True)
    detail["상태"] = detail["global_status"].apply(status_label)
    st.dataframe(detail[["상태", "supplier", "item", "material_color", "unit", "usage_per_unit", "po_required_qty", "global_secured_qty", "global_shortage_qty", "received_qty", "open_order_qty", "eta", "action"]].rename(columns={
        "supplier":"Supplier", "item":"자재명", "material_color":"Color", "unit":"단위", "usage_per_unit":"1개당 소요량", "po_required_qty":"해당 PO 필요량", "global_secured_qty":"전체 확보량", "global_shortage_qty":"전체 부족량", "received_qty":"입고 수량", "open_order_qty":"미입고 발주잔량", "eta":"ETA", "action":"요청 액션"
    }), use_container_width=True, hide_index=True)


def show_date_view(events: pd.DataFrame, purchase: pd.DataFrame):
    st.header("날짜별 보기")
    st.caption("전체 날짜 흐름은 건수 기준으로 보고, 특정 날짜나 자재를 선택하면 상세 수량을 확인합니다.")
    if events.empty:
        st.info("날짜 이벤트 데이터가 없습니다.")
        return
    tab1, tab2 = st.tabs(["전체 시계열", "자재별 날짜 흐름"])
    with tab1:
        daily = events.groupby(["date", "event_type"], as_index=False).agg(건수=("detail", "count"))
        fig = px.bar(daily, x="date", y="건수", color="event_type", barmode="group", text="건수")
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=20, b=10), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
        dates = sorted(events["date"].dt.strftime("%Y-%m-%d").unique().tolist())
        selected_date = st.selectbox("날짜 선택", dates)
        st.dataframe(events[events["date"].dt.strftime("%Y-%m-%d") == selected_date].sort_values(["event_type", "item"]), use_container_width=True, hide_index=True)
    with tab2:
        items = sorted([x for x in purchase["item"].dropna().unique().tolist() if x])
        selected_item = st.selectbox("자재 선택", items)
        ev = events[events["item"].eq(selected_item)].copy()
        if ev.empty:
            st.info("선택 자재의 날짜 이벤트가 없습니다.")
        else:
            fig = px.scatter(ev, x="date", y="event_type", size="qty", color="event_type", hover_data=["supplier", "qty", "unit", "detail"])
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10), legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(ev.sort_values("date"), use_container_width=True, hide_index=True)


def show_supplier_view(purchase: pd.DataFrame):
    st.header("공급업체별 보기")
    if purchase.empty:
        st.warning("구매데이터가 없습니다.")
        return
    data = purchase.copy()
    data["supplier"] = data["supplier"].replace("", "공급업체 확인")
    sup = data.groupby("supplier", as_index=False).agg(
        자재수=("item", "count"), 부족=("status", lambda x: int((x == "부족").sum())), 입고대기=("status", lambda x: int((x == "입고대기").sum())), 일부입고=("status", lambda x: int((x == "일부입고").sum())), 확보됨=("status", lambda x: int((x == "확보됨").sum())), 부족수량=("shortage_qty", "sum"), 부족금액=("estimated_shortage_amount", "sum"), 잔여입고확인량=("arrival_gap_qty", "sum")
    )
    sup["조치필요"] = sup["부족"] + sup["입고대기"] + sup["일부입고"]
    sup = sup.sort_values(["조치필요", "부족금액"], ascending=[False, False])
    left, right = st.columns([1, 1])
    with left:
        fig = px.bar(sup.head(12), x="supplier", y="조치필요", text="조치필요")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = px.bar(sup.head(12), x="supplier", y="부족금액", text="부족금액")
        fig.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(sup, use_container_width=True, hide_index=True)


def show_action_list(purchase: pd.DataFrame):
    st.header("액션 리스트")
    data = purchase[purchase["status"].isin(["부족", "입고대기", "일부입고", "확인필요"])].copy()
    data["상태"] = data["status"].apply(status_label)
    st.dataframe(data[["상태", "supplier", "item", "color", "unit", "loss_required", "stock_qty", "foc_qty", "order_total_qty", "secured_qty", "balance_loss_original", "shortage_qty", "received_qty", "open_order_qty", "arrival_gap_qty", "eta", "action", "basis"]].rename(columns={
        "supplier":"Supplier", "item":"자재명", "color":"Color", "unit":"단위", "loss_required":"Loss 포함 필요량", "stock_qty":"현재 재고", "foc_qty":"F.O.C", "order_total_qty":"발주 수량", "secured_qty":"확보량", "balance_loss_original":"BALANCE LOSS", "shortage_qty":"부족량", "received_qty":"입고 수량", "open_order_qty":"미입고 발주잔량", "arrival_gap_qty":"잔여 입고 확인량", "eta":"ETA", "action":"요청 액션", "basis":"판단 기준"
    }), use_container_width=True, hide_index=True)
    st.download_button("액션 리스트 CSV 다운로드", data=data.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"), file_name="material_action_list.csv", mime="text/csv")


def show_risk_analysis(
    purchase: pd.DataFrame,
    production: pd.DataFrame,
    consumption: pd.DataFrame,
    freight: pd.DataFrame,
    freight_source: str = "기본값",
):
    st.header("리스크 분석")
    st.caption("추가 물량 발생 시 공급/생산/운송 3개 측면의 리스크를 한 화면에서 확인합니다.")
    if consumption.empty or purchase.empty:
        st.warning("자재 소요량/구매 데이터가 비어 있어 리스크 분석을 진행할 수 없습니다.")
        return

    style_options = sorted(consumption["style"].dropna().unique().tolist())
    if not style_options:
        st.warning("자재 소요량 정보에 사용 가능한 STYLE이 없습니다.")
        return

    freight_available = freight is not None and not freight.empty
    dest_options = sorted(freight["destination"].dropna().unique().tolist()) if freight_available else []

    # ============== 입력 ==============
    st.subheader("입력 조건")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        default_due = (pd.Timestamp.today() + pd.Timedelta(days=45)).date()
        due_date = st.date_input("납기", value=default_due)
    with c2:
        selected_style = st.selectbox("품목 (STYLE)", style_options)
    with c3:
        add_qty = st.number_input("추가 물량 (PCS)", min_value=0.0, value=500.0, step=50.0)
    with c4:
        if freight_available:
            destination = st.selectbox("배송지", dest_options + ["기타"])
            if destination == "기타":
                destination = st.text_input("배송지 직접 입력", value="").strip().upper()
        else:
            destination = st.text_input(
                "배송지",
                value="",
                help="운송요율 데이터가 입력되지 않아 운송 리스크 계산은 비활성 상태입니다.",
            ).strip().upper()

    with st.expander("계산 파라미터 (기본값 사용 가능)", expanded=False):
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            daily_capacity = st.number_input("일일 생산능력 (PCS/day)", min_value=1.0, value=300.0, step=50.0)
        with p2:
            weight_per_unit = st.number_input("개당 무게 (kg/PCS)", min_value=0.01, value=1.5, step=0.1)
        with p3:
            critical_top_n = st.number_input("주공정 자재 TOP N (리드타임 기준)", min_value=1, max_value=20, value=5, step=1)
        with p4:
            line_change_penalty_days = st.number_input(
                "라인체인지 선택 시 추가 리드타임 (일)",
                min_value=0.0, max_value=30.0, value=0.5, step=0.5,
                help="라인 체인지 체크박스가 켜져 있을 때 생산 리드타임에 가산되는 일수입니다.",
            )

    # 라인체인지 체크박스 (선입력)
    line_change = st.checkbox(
        f"라인 체인지 발생 (체크 시 생산 리드타임 +{line_change_penalty_days:g}일 추가)",
        value=False,
        help="기존 생산 라인을 중단하고 추가 물량을 위해 라인을 변경해야 하는 경우 체크. 체크 시 위 파라미터의 가산 일수가 적용됩니다.",
    )

    due_ts = pd.Timestamp(due_date)
    supply = compute_supply_risk(selected_style, add_qty, consumption, purchase, critical_top_n=int(critical_top_n))
    prod_risk = compute_production_risk(
        selected_style, add_qty, production, daily_capacity,
        line_change=line_change, line_change_penalty_days=float(line_change_penalty_days),
    )
    trans = compute_transport_risk(due_ts, destination, add_qty, weight_per_unit, supply, prod_risk, freight)

    st.divider()

    # ============== 최상단: 총 리드타임 구성 ==============
    st.markdown("### 총 리드타임 구성")
    mat_days = float(trans["material_leadtime_days"])
    prod_days = float(trans["production_leadtime_days"])
    sea_opts = trans["options"][trans["options"]["mode"] == "SEA"] if not trans["options"].empty else pd.DataFrame()
    air_opts = trans["options"][trans["options"]["mode"] == "AIR"] if not trans["options"].empty else pd.DataFrame()
    has_sea = not sea_opts.empty
    has_air = not air_opts.empty
    sea_ship_days = float(sea_opts.iloc[0]["leadtime_days"]) if has_sea else None
    air_ship_days = float(air_opts.iloc[0]["leadtime_days"]) if has_air else None
    days_remaining = trans["days_remaining"]

    tl1, tl2, tl3, tl4, tl5, tl6 = st.columns(6)
    tl1.metric("자재 조달", f"{mat_days:.0f}일")
    tl2.metric("생산", f"{prod_days:.1f}일", help=f"기본 {prod_risk['base_production_days']:.2f}일 + 라인체인지 {prod_risk['line_change_penalty_days']:.1f}일")
    tl3.metric("운송 SEA", f"{sea_ship_days:.0f}일" if has_sea else "데이터 없음")
    tl4.metric("운송 AIR", f"{air_ship_days:.0f}일" if has_air else "데이터 없음")
    if has_sea:
        total_sea = mat_days + prod_days + sea_ship_days
        tl5.metric("총 SEA", f"{total_sea:.1f}일", delta=f"납기 여유 {days_remaining - total_sea:+.1f}일")
    else:
        tl5.metric("총 SEA", "—")
    if has_air:
        total_air = mat_days + prod_days + air_ship_days
        tl6.metric("총 AIR", f"{total_air:.1f}일", delta=f"납기 여유 {days_remaining - total_air:+.1f}일")
    else:
        tl6.metric("총 AIR", "—")

    if has_sea or has_air:
        breakdown_rows = []
        if has_sea:
            breakdown_rows += [
                {"운송수단": "SEA", "단계": "자재 조달", "일수": mat_days},
                {"운송수단": "SEA", "단계": "생산", "일수": prod_days},
                {"운송수단": "SEA", "단계": "운송", "일수": sea_ship_days},
            ]
        if has_air:
            breakdown_rows += [
                {"운송수단": "AIR", "단계": "자재 조달", "일수": mat_days},
                {"운송수단": "AIR", "단계": "생산", "일수": prod_days},
                {"운송수단": "AIR", "단계": "운송", "일수": air_ship_days},
            ]
        breakdown = pd.DataFrame(breakdown_rows)
        fig = px.bar(
            breakdown, x="일수", y="운송수단", color="단계", orientation="h", text="일수",
            color_discrete_map={"자재 조달": "#6366f1", "생산": "#f59e0b", "운송": "#10b981"},
        )
        fig.add_vline(
            x=days_remaining, line_dash="dash", line_color="#ef4444",
            annotation_text=f"납기까지 {days_remaining}일", annotation_position="top",
        )
        fig.update_traces(texttemplate="%{text:.1f}일", textposition="inside")
        fig.update_layout(height=240, margin=dict(l=10, r=10, t=30, b=10), legend_title_text="", title="단계별 리드타임 누적 (SEA vs AIR)")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info(
            "운송요율 데이터가 입력되지 않아 운송 리드타임을 계산할 수 없습니다. "
            "총 리드타임은 **자재 조달 + 생산**만 표시됩니다 — 운송요율 시트가 추가되면 자동으로 SEA/AIR이 합산됩니다."
        )
        st.metric("자재+생산 소요", f"{mat_days + prod_days:.1f}일", delta=f"납기 여유 {days_remaining - (mat_days + prod_days):+.1f}일")

    st.divider()

    # ============== 요약 메트릭 ==============
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("부족 자재 수", f"{supply['shortage_count']:,}")
    s2.metric("자재 추가 주문 비용", fmt_money(supply["total_additional_cost"]))
    s3.metric("생산 리드타임", f"{prod_risk['production_leadtime_days']:.1f}일")
    if trans["has_data"]:
        s4.metric("권장 운송 / AIR 필요", f"{trans['recommended_mode']} / {'예' if trans['air_required'] else '아니오'}")
    else:
        s4.metric("권장 운송 / AIR 필요", "데이터 없음")

    st.divider()

    # ============== 섹터 1: 공급 리스크 ==============
    st.markdown("### 1. 공급 리스크")
    if supply["materials"].empty:
        st.success("선택한 품목과 추가 물량으로 인한 자재 부족이 없습니다.")
    else:
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("부족 자재 종류", f"{supply['shortage_count']:,}")
        sc2.metric("추가 주문 총 비용", fmt_money(supply["total_additional_cost"]))
        sc3.metric("최장 리드타임", f"{supply['max_leadtime_days']}일")

        st.markdown("**부족 자재 리스트 (자재명 · 부족량 · 리드타임 · 공급업체)**")
        view = supply["materials"][[
            "item", "supplier", "new_shortage_qty", "unit", "leadtime", "leadtime_days",
            "unit_price", "additional_cost",
        ]].rename(columns={
            "item": "자재명",
            "supplier": "공급업체",
            "new_shortage_qty": "부족량",
            "unit": "단위",
            "leadtime": "리드타임",
            "leadtime_days": "리드타임(일)",
            "unit_price": "단가($)",
            "additional_cost": "추가 주문 비용($)",
        })
        st.dataframe(view, use_container_width=True, hide_index=True)

        st.markdown(f"**주공정 자재 TOP {int(critical_top_n)} (리드타임 기준)**")
        if supply["critical_materials"].empty:
            st.info("주공정으로 분류할 부족 자재가 없습니다.")
        else:
            cv = supply["critical_materials"][[
                "item", "supplier", "new_shortage_qty", "unit", "leadtime_days", "additional_cost",
            ]].rename(columns={
                "item": "자재명",
                "supplier": "공급업체",
                "new_shortage_qty": "부족량",
                "unit": "단위",
                "leadtime_days": "리드타임(일)",
                "additional_cost": "추가 주문 비용($)",
            })
            st.dataframe(cv, use_container_width=True, hide_index=True)
            fig = px.bar(
                supply["critical_materials"],
                x="leadtime_days", y="item", orientation="h",
                text="leadtime_days",
                hover_data=["supplier", "new_shortage_qty", "additional_cost"],
                labels={"leadtime_days": "리드타임(일)", "item": "자재명"},
            )
            fig.update_traces(texttemplate="%{text}일", textposition="outside")
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ============== 섹터 2: 생산 리스크 ==============
    st.markdown("### 2. 생산 리스크")
    pc1, pc2, pc3, pc4 = st.columns(4)
    pc1.metric("기본 생산일수", f"{prod_risk['base_production_days']:.2f}일")
    pc2.metric("라인 체인지 가산", f"+{prod_risk['line_change_penalty_days']:.1f}일")
    pc3.metric("생산 리드타임 합계", f"{prod_risk['production_leadtime_days']:.1f}일")
    pc4.metric("일일 생산능력", f"{int(prod_risk['daily_capacity']):,} PCS/day")

    if line_change:
        st.warning(f"라인 체인지 발생 — 생산 리드타임에 {prod_risk['line_change_penalty_days']:.1f}일이 가산되었습니다.")
    else:
        st.info("라인 체인지 미발생 — 가산 시간 없음.")

    # 참고용 자동감지
    if prod_risk["line_change_detected"]:
        st.caption(f"참고: 선택 품목 '{selected_style}'은 현재 생산계획에 존재하지 않습니다 (신규 STYLE).")
    else:
        st.caption(f"참고: 선택 품목 '{selected_style}'은 현재 생산계획에 존재합니다 (기존 STYLE).")

    if not prod_risk["existing_runs"].empty:
        st.markdown("**기존 생산계획 (해당 STYLE)**")
        runs = prod_risk["existing_runs"][[
            "buyer_po", "color", "prod_qty", "sewing_from", "sewing_to", "ship_date", "material_status",
        ]].rename(columns={
            "buyer_po": "PO",
            "color": "Color",
            "prod_qty": "생산수량",
            "sewing_from": "Sewing 시작",
            "sewing_to": "Sewing 종료",
            "ship_date": "Ship Date",
            "material_status": "자재 상태",
        })
        st.dataframe(runs, use_container_width=True, hide_index=True)

    st.divider()

    # ============== 섹터 3: 운송 리스크 ==============
    st.markdown("### 3. 운송 리스크")
    st.caption(f"운송요율 데이터 출처: **{freight_source}**")
    if not trans["has_data"]:
        st.warning(
            "운송 리스크 계산을 위한 운송요율 데이터가 없습니다. "
            "엑셀에 **'운송요율' 시트**를 추가하면 자동으로 계산됩니다."
        )
        with st.expander("운송요율 시트 형식 안내"):
            st.markdown(
                """
                | 배송지 (destination) | 운송수단 (mode) | 리드타임(일) (leadtime_days) | 단가($/kg) (rate_usd_per_kg) |
                |---|---|---|---|
                | USA | SEA | 30 | 3.5 |
                | USA | AIR | 5  | 12.0 |
                | EU  | SEA | 35 | 4.0 |
                | …   | …   | …  | …   |

                - 컬럼명은 한글/영문 모두 인식합니다.
                - `mode`는 SEA / AIR (또는 해상 / 항공)로 입력해 주세요.
                """
            )
        return

    tc1, tc2, tc3, tc4 = st.columns(4)
    tc1.metric("납기까지 남은 일수", f"{trans['days_remaining']}일")
    tc2.metric("운송 전 소요 (자재+생산)", f"{trans['material_leadtime_days'] + trans['production_leadtime_days']:.1f}일")
    tc3.metric("총 운송 무게", f"{trans['total_weight_kg']:,.0f} kg")
    tc4.metric("AIR 필요 여부", "예" if trans["air_required"] else "아니오")

    if trans["options"].empty:
        st.warning(f"배송지 '{destination}'에 대한 운송요율 정보를 찾을 수 없습니다.")
    else:
        st.markdown("**SEA vs AIR 비교**")
        opts_view = trans["options"][[
            "mode", "leadtime_days", "rate_usd_per_kg", "total_weight_kg",
            "transport_cost", "total_leadtime_days", "margin_days", "meets_deadline",
        ]].rename(columns={
            "mode": "운송수단",
            "leadtime_days": "운송 리드타임(일)",
            "rate_usd_per_kg": "단가($/kg)",
            "total_weight_kg": "무게(kg)",
            "transport_cost": "운송 비용($)",
            "total_leadtime_days": "총 소요(일)",
            "margin_days": "납기 여유(일)",
            "meets_deadline": "납기 충족",
        })
        st.dataframe(opts_view, use_container_width=True, hide_index=True)

        comp_df = trans["options"].copy()
        cc1, cc2 = st.columns(2)
        with cc1:
            fig = px.bar(comp_df, x="mode", y="transport_cost", color="mode", text="transport_cost",
                         color_discrete_map={"SEA": "#10b981", "AIR": "#ef4444"},
                         labels={"mode": "운송수단", "transport_cost": "운송 비용($)"})
            fig.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=30, b=10), title="운송 비용 비교")
            st.plotly_chart(fig, use_container_width=True)
        with cc2:
            fig = px.bar(comp_df, x="mode", y="total_leadtime_days", color="mode", text="total_leadtime_days",
                         color_discrete_map={"SEA": "#10b981", "AIR": "#ef4444"},
                         labels={"mode": "운송수단", "total_leadtime_days": "총 소요(일)"})
            fig.add_hline(y=trans["days_remaining"], line_dash="dash", line_color="#6b7280",
                          annotation_text=f"납기까지 {trans['days_remaining']}일", annotation_position="top right")
            fig.update_traces(texttemplate="%{text:.1f}일", textposition="outside")
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=30, b=10), title="납기 vs 총 리드타임")
            st.plotly_chart(fig, use_container_width=True)

        if trans["air_required"]:
            air_row = comp_df[comp_df["mode"] == "AIR"]
            sea_row = comp_df[comp_df["mode"] == "SEA"]
            if not air_row.empty:
                msg = f"**SEA로는 납기를 맞출 수 없습니다.** AIR 운송 필요 — 비용 {fmt_money(air_row.iloc[0]['transport_cost'])}, 리드타임 {int(air_row.iloc[0]['leadtime_days'])}일."
                if not sea_row.empty:
                    diff = float(air_row.iloc[0]["transport_cost"] - sea_row.iloc[0]["transport_cost"])
                    msg += f" SEA 대비 추가 비용 +{fmt_money(diff)}."
                st.error(msg)
        else:
            st.success(f"SEA 운송으로 납기 충족 가능합니다. (권장: {trans['recommended_mode']})")


def show_quality(purchase: pd.DataFrame, production: pd.DataFrame, consumption: pd.DataFrame, inbound: pd.DataFrame, outbound: pd.DataFrame):
    st.header("데이터 점검")
    q = pd.DataFrame([
        {"데이터":"구매데이터", "행 수":len(purchase), "상태":"OK" if len(purchase) else "확인 필요"},
        {"데이터":"생산계획", "행 수":len(production), "상태":"OK" if len(production) else "확인 필요"},
        {"데이터":"자재 소요량 정보", "행 수":len(consumption), "상태":"OK" if len(consumption) else "확인 필요"},
        {"데이터":"자재입고내역", "행 수":len(inbound), "상태":"OK" if len(inbound) else "선택"},
        {"데이터":"자재출고내역", "행 수":len(outbound), "상태":"OK" if len(outbound) else "선택"},
    ])
    st.dataframe(q, use_container_width=True, hide_index=True)
    if not purchase.empty:
        max_abs = purchase["balance_loss_diff"].abs().max()
        bad = purchase[purchase["balance_loss_diff"].abs() > 1e-6]
        st.metric("BALANCE LOSS 원본-앱 계산 최대 차이", f"{max_abs:,.6f}")
        if bad.empty:
            st.success("구매데이터의 BALANCE LOSS 수식과 앱 계산값이 일치합니다.")
        else:
            st.warning(f"원본 수식과 앱 계산값 차이 발생 행: {len(bad)}개")
            st.dataframe(bad[["source_row", "item", "balance_loss_original", "balance_loss_calc", "balance_loss_diff"]], use_container_width=True, hide_index=True)
    with st.expander("파싱된 구매데이터"):
        st.dataframe(purchase.head(300), use_container_width=True)
    with st.expander("파싱된 생산계획"):
        st.dataframe(production.head(300), use_container_width=True)
    with st.expander("파싱된 자재 소요량 정보"):
        st.dataframe(consumption.head(300), use_container_width=True)
    with st.expander("파싱된 입고내역"):
        st.dataframe(inbound.head(300), use_container_width=True)
    with st.expander("파싱된 출고내역"):
        st.dataframe(outbound.head(300), use_container_width=True)


# =========================================================
# 5. Main
# =========================================================

def main():
    st.sidebar.title("설정")
    uploaded = st.sidebar.file_uploader("엑셀 업로드", type=["xlsx", "xlsm"])
    tolerance_qty = st.sidebar.number_input("미세 부족 허용 수량", min_value=0.0, max_value=100.0, value=1.0, step=1.0)
    if uploaded is None:
        show_intro()
        return
    try:
        sheet_names, sheets = load_excel(uploaded.getvalue())
        sh_purchase = find_sheet(sheet_names, ["구매데이터"])
        sh_prod = find_sheet(sheet_names, ["생산계획"])
        sh_in = find_sheet(sheet_names, ["자재입고"])
        sh_out = find_sheet(sheet_names, ["자재출고"])
        sh_freight = find_sheet(sheet_names, ["운송요율", "운송", "freight", "Freight"])
        raw_purchase = sheets.get(sh_purchase, pd.DataFrame())
        purchase = parse_purchase(raw_purchase, tolerance_qty=float(tolerance_qty))
        consumption = parse_consumption(raw_purchase)
        production = parse_production(sheets.get(sh_prod, pd.DataFrame()))
        inbound = parse_flow(sheets.get(sh_in, pd.DataFrame()), "입고")
        outbound = parse_flow(sheets.get(sh_out, pd.DataFrame()), "출고")
        if sh_freight:
            freight = parse_freight(sheets.get(sh_freight, pd.DataFrame()))
            if freight.empty:
                freight_source = f"엑셀 '{sh_freight}' 시트 — 인식되었으나 형식 불일치 또는 데이터 없음"
            else:
                freight_source = f"엑셀 '{sh_freight}' 시트 ({len(freight)}행)"
        else:
            freight = EMPTY_FREIGHT.copy()
            freight_source = "데이터 없음 — 엑셀에 '운송요율' 시트가 없습니다 (운송 리스크 계산 비활성)"
        events = build_events(purchase, production, inbound, outbound)
    except Exception as e:
        st.error("엑셀 분석 중 오류가 발생했습니다.")
        st.exception(e)
        return

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "전체 대시보드", "자재별", "PO별", "날짜별", "공급업체별", "액션 리스트", "리스크 분석", "데이터 점검"
    ])
    with tab1:
        show_dashboard(purchase, production)
    with tab2:
        show_material_view(purchase)
    with tab3:
        show_po_view(production, consumption, purchase)
    with tab4:
        show_date_view(events, purchase)
    with tab5:
        show_supplier_view(purchase)
    with tab6:
        show_action_list(purchase)
    with tab7:
        show_risk_analysis(purchase, production, consumption, freight, freight_source=freight_source)
    with tab8:
        show_quality(purchase, production, consumption, inbound, outbound)


if __name__ == "__main__":
    main()
