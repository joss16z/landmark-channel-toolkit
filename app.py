from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any, Optional

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

st.set_page_config(page_title="Landmark Channel Toolkit", page_icon="🏢", layout="wide")

# -----------------------------
# Helpers
# -----------------------------

def clean_unit(value: Any) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().upper()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def clean_area(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    s = str(value).strip().replace("㎡", "").replace("sqm", "").replace("SQM", "").replace(",", "")
    try:
        n = float(s)
        return str(int(n)) if n.is_integer() else str(n)
    except Exception:
        return s


def money_number(value: Any) -> Optional[float]:
    if value is None or pd.isna(value):
        return None
    s = str(value).strip()
    if not s or not re.search(r"\d", s):
        return None
    raw = re.sub(r"[^0-9.\-]", "", s)
    if raw in {"", "-"}:
        return None
    try:
        return float(raw)
    except Exception:
        return None


def format_money(value: Any) -> str:
    n = money_number(value)
    if n is None:
        return "" if value is None or pd.isna(value) else str(value).replace("A$", "$")
    return f"${n:,.0f}"


def extract_money_text(text: str) -> str:
    m = re.search(r"\$\s?[0-9,]+", text or "")
    return m.group(0).replace(" ", "") if m else ""


def normalize_price_text(text: str) -> str:
    n = money_number(text)
    return format_money(n) if n is not None else ""


def find_unit_in_text(text: str) -> Optional[str]:
    m = re.search(r"\bUnit\s+([A-Z]*\d+[A-Z]*)\b", text or "", flags=re.I)
    return clean_unit(m.group(1)) if m else None


def find_units_in_text(text: str) -> list[str]:
    return [clean_unit(x) for x in re.findall(r"\bUnit\s+([A-Z]*\d+[A-Z]*)\b", text or "", flags=re.I)]


def best_match(columns, candidates):
    lower = {str(c).lower().strip(): c for c in columns}
    for cand in candidates:
        if cand.lower().strip() in lower:
            return lower[cand.lower().strip()]
    # fuzzy contains
    for cand in candidates:
        key = cand.lower().strip()
        for c in columns:
            if key in str(c).lower().strip():
                return c
    return None


def index_or_zero(columns, value):
    return columns.index(value) if value in columns else 0


def set_shape_text_keep_basic_format(shape, new_text: str) -> None:
    tf = shape.text_frame
    font_name = font_size = bold = italic = font_color = None
    align = PP_ALIGN.CENTER
    try:
        p0 = tf.paragraphs[0]
        align = p0.alignment or PP_ALIGN.CENTER
        if p0.runs:
            r0 = p0.runs[0]
            font_name = r0.font.name
            font_size = r0.font.size
            bold = r0.font.bold
            italic = r0.font.italic
            if r0.font.color and r0.font.color.rgb:
                font_color = r0.font.color.rgb
    except Exception:
        pass

    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, line in enumerate(str(new_text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        for run in p.runs:
            if font_name:
                run.font.name = font_name
            if font_size:
                run.font.size = font_size
            if bold is not None:
                run.font.bold = bold
            if italic is not None:
                run.font.italic = italic
            if font_color:
                run.font.color.rgb = font_color


def mark_shape_red(shape):
    try:
        shape.line.color.rgb = RGBColor(255, 0, 0)
        shape.line.width = 25400  # about 2pt
    except Exception:
        pass


@dataclass
class ExcelConfig:
    unit_col: str
    price_col: str
    status_col: Optional[str] = None
    internal_col: Optional[str] = None
    external_col: Optional[str] = None
    net_price_col: Optional[str] = None
    available_only: bool = True


def load_excel(uploaded_file, sheet_name: str) -> pd.DataFrame:
    data = uploaded_file.read()
    uploaded_file.seek(0)
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)


def build_price_table(df: pd.DataFrame, cfg: ExcelConfig) -> pd.DataFrame:
    work = df.copy()
    work["__unit"] = work[cfg.unit_col].apply(clean_unit)
    work["__excel_price"] = work[cfg.price_col].apply(format_money)
    work["__price_num"] = work[cfg.price_col].apply(money_number)
    if cfg.status_col:
        work["__status"] = work[cfg.status_col].fillna("").astype(str).str.strip()
    else:
        work["__status"] = ""

    if cfg.available_only and cfg.status_col:
        work = work[work["__status"].str.lower().eq("available")]

    work = work[work["__unit"].astype(bool)].drop_duplicates("__unit", keep="first")
    return work


def update_unit_price_line(text: str, unit: str, excel_price: str) -> tuple[str, bool, str]:
    """Update `Unit XXX - $...` inside text. Returns new_text, changed, ppt_price."""
    ppt_price = ""
    changed = False

    # Match Unit 204 - $1,083,000, allowing odd spacing.
    pattern = re.compile(rf"(\bUnit\s+{re.escape(unit)}\b\s*[-–—]\s*)(\$\s?[0-9,]+)", flags=re.I)

    def repl(m):
        nonlocal ppt_price, changed
        ppt_price = m.group(2).replace(" ", "")
        if normalize_price_text(ppt_price) != normalize_price_text(excel_price):
            changed = True
        return f"Unit {unit} - {excel_price}"

    new_text, n = pattern.subn(repl, text)
    if n:
        return new_text, changed, ppt_price

    # Fallback: if line has Unit XXX but no price pattern, append price.
    unit_pattern = re.compile(rf"\bUnit\s+{re.escape(unit)}\b", flags=re.I)
    if unit_pattern.search(text):
        existing = extract_money_text(text)
        ppt_price = existing
        if normalize_price_text(existing) != normalize_price_text(excel_price):
            changed = True
        new_text = unit_pattern.sub(f"Unit {unit} - {excel_price}", text, count=1)
        return new_text, changed, ppt_price

    return text, False, ppt_price


def run_price_update(ppt_bytes: bytes, price_df: pd.DataFrame, mark_extra: bool = True):
    prs = Presentation(io.BytesIO(ppt_bytes))
    price_map = {r["__unit"]: r["__excel_price"] for _, r in price_df.iterrows()}
    excel_units = set(price_map.keys())
    ppt_units = set()
    rows = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            if not getattr(shape, "has_text_frame", False):
                continue
            old_text = shape.text_frame.text or ""
            units = find_units_in_text(old_text)
            if not units:
                continue

            new_text = old_text
            shape_changed = False
            shape_has_extra = False

            for unit in units:
                ppt_units.add(unit)
                old_price = extract_money_text(new_text)
                if unit in price_map:
                    updated_text, price_changed, ppt_price = update_unit_price_line(new_text, unit, price_map[unit])
                    new_text = updated_text
                    result = "Price Updated" if price_changed else "Matched / Same Price"
                    shape_changed = shape_changed or price_changed or (updated_text != old_text)
                    rows.append({
                        "Unit": unit,
                        "Slide": slide_idx,
                        "Shape": shape_idx,
                        "PPT Price": ppt_price or old_price,
                        "Excel Price": price_map[unit],
                        "Difference": "" if not price_changed else "Check",
                        "Result": result,
                        "Action": "Update PPT price" if price_changed else "No action",
                    })
                else:
                    shape_has_extra = True
                    rows.append({
                        "Unit": unit,
                        "Slide": slide_idx,
                        "Shape": shape_idx,
                        "PPT Price": extract_money_text(new_text),
                        "Excel Price": "",
                        "Difference": "",
                        "Result": "Extra in PPT",
                        "Action": "Review / remove or mark unavailable",
                    })

            if new_text != old_text:
                set_shape_text_keep_basic_format(shape, new_text)
            if mark_extra and shape_has_extra:
                mark_shape_red(shape)

    for unit in sorted(excel_units - ppt_units):
        rows.append({
            "Unit": unit,
            "Slide": "",
            "Shape": "",
            "PPT Price": "",
            "Excel Price": price_map[unit],
            "Difference": "",
            "Result": "Missing in PPT",
            "Action": "Add label to PPT",
        })

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    report = pd.DataFrame(rows)
    if not report.empty:
        order = {
            "Price Updated": 1,
            "Missing in PPT": 2,
            "Extra in PPT": 3,
            "Matched / Same Price": 4,
        }
        report["__order"] = report["Result"].map(order).fillna(99)
        report = report.sort_values(["__order", "Unit"], kind="stable").drop(columns="__order")
    return out.getvalue(), report


def report_to_xlsx(report: pd.DataFrame, price_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary = report["Result"].value_counts().rename_axis("Result").reset_index(name="Count") if not report.empty else pd.DataFrame(columns=["Result", "Count"])
        summary.loc[len(summary)] = ["Excel Available Units", len(price_df)]
        summary.to_excel(writer, index=False, sheet_name="Summary")
        report.to_excel(writer, index=False, sheet_name="All Checks")
        for name, result in [
            ("Price Changed", "Price Updated"),
            ("Missing in PPT", "Missing in PPT"),
            ("Extra in PPT", "Extra in PPT"),
            ("Matched", "Matched / Same Price"),
        ]:
            subset = report[report["Result"].eq(result)] if not report.empty else pd.DataFrame()
            subset.to_excel(writer, index=False, sheet_name=name[:31])
    output.seek(0)
    return output.getvalue()

# -----------------------------
# UI
# -----------------------------

st.sidebar.title("功能")
page = st.sidebar.radio("选择功能", ["✅ PPT价格更新", "📋 PPT核对Report", "🏗️ Floorplate Generator", "📊 Price Compare", "📦 Agent Package Generator"])

st.title("🏢 Landmark Channel Toolkit")
st.caption("Floorplate price update / audit tools")

if page == "✅ PPT价格更新":
    st.header("✅ 根据 Excel 自动更新 PPT 里的 Unit 价格")
    st.write("匹配 PPT 中的 `Unit xxx - $xxx`，用 Excel 的价格替换。可同时生成检查报告。")

    c1, c2 = st.columns(2)
    with c1:
        ppt_file = st.file_uploader("Upload Floorplate PPT / PPTX", type=["pptx", "ppt"], key="price_ppt")
    with c2:
        excel_file = st.file_uploader("Upload Price List Excel", type=["xls", "xlsx"], key="price_excel")

    if excel_file:
        xls = pd.ExcelFile(excel_file)
        sheet_name = st.selectbox("Excel sheet", xls.sheet_names, index=0, key="price_sheet")
        df = load_excel(excel_file, sheet_name)
        columns = list(df.columns)

        default_unit = best_match(columns, ["Unit Number", "Unit", "Apt #", "Apt", "Apartment", "Unit No"])
        default_price = best_match(columns, ["Contract Price", "Price", "List Price"])
        default_status = best_match(columns, ["Status", "Availability"])

        c1, c2, c3 = st.columns(3)
        with c1:
            unit_col = st.selectbox("Unit column", columns, index=index_or_zero(columns, default_unit), key="price_unit_col")
        with c2:
            price_col = st.selectbox("Price column", columns, index=index_or_zero(columns, default_price), key="price_price_col")
        with c3:
            status_col = st.selectbox("Status column", [None] + columns, index=([None] + columns).index(default_status) if default_status else 0, key="price_status_col")

        c4, c5 = st.columns(2)
        with c4:
            available_only = st.checkbox("Only use Available units from Excel", value=True)
        with c5:
            mark_extra = st.checkbox("Mark PPT extra units with red border", value=True)

        st.caption("Excel preview")
        st.dataframe(df.head(10), use_container_width=True)

        if ppt_file and st.button("Update PPT Prices", type="primary"):
            cfg = ExcelConfig(unit_col=unit_col, price_col=price_col, status_col=status_col, available_only=available_only)
            price_df = build_price_table(df, cfg)
            updated_ppt, report = run_price_update(ppt_file.read(), price_df, mark_extra=mark_extra)
            report_xlsx = report_to_xlsx(report, price_df)

            st.success("Done. PPT prices updated and report generated.")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Excel Available", len(price_df))
            m2.metric("Price Updated", int((report["Result"] == "Price Updated").sum()) if not report.empty else 0)
            m3.metric("Missing in PPT", int((report["Result"] == "Missing in PPT").sum()) if not report.empty else 0)
            m4.metric("Extra in PPT", int((report["Result"] == "Extra in PPT").sum()) if not report.empty else 0)

            st.download_button("Download Updated PPT", updated_ppt, file_name="Updated_Floorplate_Prices.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            st.download_button("Download Audit Report", report_xlsx, file_name="Floorplate_Price_Update_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(report, use_container_width=True)
    else:
        st.info("Upload a Price List Excel to start.")

elif page == "📋 PPT核对Report":
    st.header("📋 PPT 核对 Report")
    st.info("当前版本可在“PPT价格更新”中同时生成 Report。后续可把 Audit-only 独立增强。")

else:
    st.header(page)
    st.info("This module is planned. The price update module is ready now.")
