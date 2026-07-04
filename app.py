from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any, Optional

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor

st.set_page_config(page_title="Landmark Channel Toolkit", page_icon="🏢", layout="wide")

# -----------------------------
# Basic helpers
# -----------------------------

def clean_unit(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    s = str(value).strip().upper()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def money_number(value: Any) -> Optional[float]:
    if value is None or pd.isna(value):
        return None
    s = str(value).strip()
    if not s or not re.search(r"\d", s):
        return None
    raw = re.sub(r"[^0-9.\-]", "", s)
    if raw in {"", "-", "."}:
        return None
    try:
        return float(raw)
    except Exception:
        return None


def format_money(value: Any) -> str:
    n = money_number(value)
    if n is None:
        if value is None or pd.isna(value):
            return ""
        return str(value).replace("A$", "$").strip()
    return f"${n:,.0f}"


def normalize_money(value: Any) -> str:
    n = money_number(value)
    return format_money(n) if n is not None else ""


def best_match(columns, candidates):
    lower = {str(c).lower().strip(): c for c in columns}
    for cand in candidates:
        key = cand.lower().strip()
        if key in lower:
            return lower[key]
    for cand in candidates:
        key = cand.lower().strip()
        for c in columns:
            if key in str(c).lower().strip():
                return c
    return None


def index_or_zero(columns, value):
    return columns.index(value) if value in columns else 0


@dataclass
class ExcelConfig:
    unit_col: str
    price_col: str
    status_col: Optional[str] = None
    available_only: bool = True


def read_uploaded_bytes(uploaded_file) -> bytes:
    data = uploaded_file.getvalue()
    return data


def get_sheet_names(uploaded_file) -> list[str]:
    data = read_uploaded_bytes(uploaded_file)
    try:
        xls = pd.ExcelFile(io.BytesIO(data))
        return xls.sheet_names
    except Exception:
        # Some Arcanite .xls exports are HTML tables saved with xls extension.
        try:
            tables = pd.read_html(io.BytesIO(data))
            return [f"Table {i+1}" for i in range(len(tables))] or ["Sheet1"]
        except Exception:
            return ["Sheet1"]


def load_excel(uploaded_file, sheet_name: str) -> pd.DataFrame:
    data = read_uploaded_bytes(uploaded_file)
    try:
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)
    except Exception as e1:
        # Fallback for HTML-as-XLS exports.
        try:
            tables = pd.read_html(io.BytesIO(data))
            if sheet_name.startswith("Table "):
                idx = int(sheet_name.split(" ")[1]) - 1
            else:
                idx = 0
            return tables[idx]
        except Exception:
            raise e1


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


# -----------------------------
# PPT parsing and update helpers
# -----------------------------

UNIT_PATTERN = re.compile(r"\bUnit\s*([A-Z]*\d+[A-Z]*)\b", flags=re.I)
PRICE_PATTERN = re.compile(r"\$\s?[0-9,]+")


def find_units_in_text(text: str) -> list[str]:
    return [clean_unit(x) for x in UNIT_PATTERN.findall(text or "")]


def extract_unit_prices_from_text(text: str) -> dict[str, str]:
    """Extract Unit -> nearest price after that Unit before the next Unit.
    Handles text like: Unit 204 - $1,083,000 and Unit4104 - $1,595,000.
    """
    text = text or ""
    matches = list(UNIT_PATTERN.finditer(text))
    result = {}
    for i, m in enumerate(matches):
        unit = clean_unit(m.group(1))
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        segment = text[start:end]
        price_m = PRICE_PATTERN.search(segment)
        result[unit] = price_m.group(0).replace(" ", "") if price_m else ""
    return result


def parse_ppt_units(prs: Presentation) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]]]:
    """Return Unit -> first occurrence info and all occurrences."""
    first = {}
    all_rows = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text_frame.text or ""
            unit_prices = extract_unit_prices_from_text(text)
            for unit, price in unit_prices.items():
                row = {
                    "Unit": unit,
                    "PPT Price": normalize_money(price),
                    "Slide": slide_idx,
                    "Shape": shape_idx,
                    "Text": text,
                }
                all_rows.append(row)
                if unit not in first:
                    first[unit] = row
    return first, all_rows


def replace_price_in_run_text(run_text: str, price_map: dict[str, str]) -> tuple[str, bool]:
    """Replace prices only inside the run text, preserving the rest of the run formatting.
    It supports multiple Unit-price pairs in one run.
    """
    original = run_text
    matches = list(UNIT_PATTERN.finditer(run_text or ""))
    if not matches:
        return run_text, False

    pieces = []
    last = 0
    changed = False
    for i, m in enumerate(matches):
        unit = clean_unit(m.group(1))
        seg_start = m.end()
        seg_end = matches[i + 1].start() if i + 1 < len(matches) else len(run_text)
        pieces.append(run_text[last:seg_start])
        segment = run_text[seg_start:seg_end]
        if unit in price_map:
            new_price = price_map[unit]
            def repl_price(pm):
                nonlocal changed
                old_price = normalize_money(pm.group(0))
                if old_price != normalize_money(new_price):
                    changed = True
                return new_price
            segment2, n = PRICE_PATTERN.subn(repl_price, segment, count=1)
            pieces.append(segment2)
        else:
            pieces.append(segment)
        last = seg_end
    pieces.append(run_text[last:])
    new_text = "".join(pieces)
    return new_text, changed or (new_text != original)


def update_ppt_prices_keep_format(ppt_bytes: bytes, price_df: pd.DataFrame) -> tuple[bytes, pd.DataFrame]:
    prs = Presentation(io.BytesIO(ppt_bytes))
    price_map = {r["__unit"]: r["__excel_price"] for _, r in price_df.iterrows()}

    before_map, occurrences = parse_ppt_units(prs)
    rows = []
    updated_units = set()
    skipped_units = set()

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text_frame.text or ""
            units = find_units_in_text(text)
            if not units:
                continue

            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    new_run_text, changed = replace_price_in_run_text(run.text, price_map)
                    if new_run_text != run.text:
                        run.text = new_run_text
                        for u in find_units_in_text(new_run_text):
                            if u in price_map:
                                updated_units.add(u)

            # Mark units with no price replacement opportunity as skipped if in Excel but no price found.
            after_text = shape.text_frame.text or ""
            after_prices = extract_unit_prices_from_text(after_text)
            before_prices = extract_unit_prices_from_text(text)
            for unit in units:
                if unit in price_map:
                    if normalize_money(after_prices.get(unit, "")) == normalize_money(price_map[unit]):
                        result = "Updated / Same after update"
                        action = "No action" if normalize_money(before_prices.get(unit, "")) == normalize_money(price_map[unit]) else "Price updated"
                    else:
                        result = "Skipped - complex text formatting"
                        action = "Manual check / make sure Unit and price are in the same text run"
                        skipped_units.add(unit)
                    rows.append({
                        "Unit": unit,
                        "Slide": slide_idx,
                        "Shape": shape_idx,
                        "PPT Price Before": normalize_money(before_prices.get(unit, "")),
                        "Excel Price": price_map[unit],
                        "PPT Price After": normalize_money(after_prices.get(unit, "")),
                        "Result": result,
                        "Action": action,
                    })
                else:
                    rows.append({
                        "Unit": unit,
                        "Slide": slide_idx,
                        "Shape": shape_idx,
                        "PPT Price Before": normalize_money(before_prices.get(unit, "")),
                        "Excel Price": "",
                        "PPT Price After": normalize_money(after_prices.get(unit, "")),
                        "Result": "Extra in PPT",
                        "Action": "Review / remove if unavailable",
                    })

    ppt_units = set(before_map.keys())
    excel_units = set(price_map.keys())
    for unit in sorted(excel_units - ppt_units):
        rows.append({
            "Unit": unit,
            "Slide": "",
            "Shape": "",
            "PPT Price Before": "",
            "Excel Price": price_map[unit],
            "PPT Price After": "",
            "Result": "Missing in PPT",
            "Action": "Add label manually if needed",
        })

    out = io.BytesIO()
    prs.save(out)
    report = pd.DataFrame(rows)
    report = order_report(report)
    return out.getvalue(), report


def audit_only(ppt_bytes: bytes, price_df: pd.DataFrame) -> pd.DataFrame:
    prs = Presentation(io.BytesIO(ppt_bytes))
    ppt_map, occurrences = parse_ppt_units(prs)
    price_map = {r["__unit"]: r["__excel_price"] for _, r in price_df.iterrows()}
    excel_units = set(price_map.keys())
    ppt_units = set(ppt_map.keys())
    rows = []

    for unit in sorted(excel_units & ppt_units):
        ppt_price = normalize_money(ppt_map[unit].get("PPT Price", ""))
        excel_price = normalize_money(price_map[unit])
        changed = ppt_price != excel_price
        rows.append({
            "Unit": unit,
            "Slide": ppt_map[unit].get("Slide", ""),
            "Shape": ppt_map[unit].get("Shape", ""),
            "PPT Price": ppt_price,
            "Excel Price": excel_price,
            "Result": "Price Changed" if changed else "Matched / Same Price",
            "Action": "Update PPT price" if changed else "No action",
        })

    for unit in sorted(excel_units - ppt_units):
        rows.append({
            "Unit": unit,
            "Slide": "",
            "Shape": "",
            "PPT Price": "",
            "Excel Price": normalize_money(price_map[unit]),
            "Result": "Missing in PPT",
            "Action": "Add label manually if needed",
        })

    for unit in sorted(ppt_units - excel_units):
        rows.append({
            "Unit": unit,
            "Slide": ppt_map[unit].get("Slide", ""),
            "Shape": ppt_map[unit].get("Shape", ""),
            "PPT Price": normalize_money(ppt_map[unit].get("PPT Price", "")),
            "Excel Price": "",
            "Result": "Extra in PPT",
            "Action": "Review / remove if unavailable",
        })

    return order_report(pd.DataFrame(rows))


def order_report(report: pd.DataFrame) -> pd.DataFrame:
    if report.empty:
        return report
    order = {
        "Price Changed": 1,
        "Skipped - complex text formatting": 2,
        "Missing in PPT": 3,
        "Extra in PPT": 4,
        "Updated / Same after update": 5,
        "Matched / Same Price": 6,
    }
    report["__order"] = report["Result"].map(order).fillna(99)
    report = report.sort_values(["__order", "Unit"], kind="stable").drop(columns="__order")
    return report


def report_to_xlsx(report: pd.DataFrame, price_df: pd.DataFrame, project_name: str = "Floorplate") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_rows = []
        summary_rows.append({"Metric": "Project", "Value": project_name})
        summary_rows.append({"Metric": "Excel Available Units", "Value": len(price_df)})
        if not report.empty:
            counts = report["Result"].value_counts().to_dict()
            for key in ["Price Changed", "Skipped - complex text formatting", "Missing in PPT", "Extra in PPT", "Matched / Same Price", "Updated / Same after update"]:
                if key in counts:
                    summary_rows.append({"Metric": key, "Value": counts[key]})
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="Summary")
        report.to_excel(writer, index=False, sheet_name="All Checks")
        sheet_map = [
            ("Price Changed", ["Price Changed"]),
            ("Skipped", ["Skipped - complex text formatting"]),
            ("Missing in PPT", ["Missing in PPT"]),
            ("Extra in PPT", ["Extra in PPT"]),
            ("Matched", ["Matched / Same Price", "Updated / Same after update"]),
        ]
        for sheet, results in sheet_map:
            subset = report[report["Result"].isin(results)] if not report.empty else pd.DataFrame()
            subset.to_excel(writer, index=False, sheet_name=sheet[:31])

        # Basic formatting
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 42)
    return output.getvalue()


# -----------------------------
# Shared UI controls
# -----------------------------

def excel_controls(key_prefix: str):
    excel_file = st.file_uploader("Upload Price List Excel", type=["xls", "xlsx"], key=f"{key_prefix}_excel")
    if not excel_file:
        return None, None, None

    sheet_names = get_sheet_names(excel_file)
    sheet_name = st.selectbox("Excel sheet", sheet_names, index=0, key=f"{key_prefix}_sheet")
    try:
        df = load_excel(excel_file, sheet_name)
    except Exception as e:
        st.error("Excel 读取失败。请尝试把文件另存为 .xlsx 后再上传。")
        st.exception(e)
        return None, None, None

    columns = list(df.columns)
    default_unit = best_match(columns, ["Unit Number", "Unit", "Apt #", "Apt", "Apartment", "Unit No"])
    default_price = best_match(columns, ["Contract Price", "Price", "List Price"])
    default_status = best_match(columns, ["Status", "Availability"])

    c1, c2, c3 = st.columns(3)
    with c1:
        unit_col = st.selectbox("Unit column", columns, index=index_or_zero(columns, default_unit), key=f"{key_prefix}_unit_col")
    with c2:
        price_col = st.selectbox("Price column", columns, index=index_or_zero(columns, default_price), key=f"{key_prefix}_price_col")
    with c3:
        status_options = [None] + columns
        status_index = status_options.index(default_status) if default_status in status_options else 0
        status_col = st.selectbox("Status column", status_options, index=status_index, key=f"{key_prefix}_status_col")

    available_only = st.checkbox("Only use Available units from Excel", value=True, key=f"{key_prefix}_available_only")
    cfg = ExcelConfig(unit_col=unit_col, price_col=price_col, status_col=status_col, available_only=available_only)
    price_df = build_price_table(df, cfg)
    st.caption("Excel preview")
    st.dataframe(df.head(10), use_container_width=True)
    return df, price_df, cfg


def show_metrics(report: pd.DataFrame, price_df: pd.DataFrame):
    counts = report["Result"].value_counts().to_dict() if not report.empty else {}
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Excel Available", len(price_df))
    m2.metric("Price Changed", counts.get("Price Changed", 0) + counts.get("Skipped - complex text formatting", 0))
    m3.metric("Missing in PPT", counts.get("Missing in PPT", 0))
    m4.metric("Extra in PPT", counts.get("Extra in PPT", 0))


# -----------------------------
# App UI
# -----------------------------

st.sidebar.title("功能")
page = st.sidebar.radio(
    "选择功能",
    ["📋 PPT核对Report", "✅ PPT价格更新", "🏗️ Floorplate Generator", "📊 Price Compare", "📦 Agent Package Generator"],
)

st.title("🏢 Landmark Channel Toolkit")
st.caption("V3 — PPT 核对 Report + PPT 价格更新")

if page == "📋 PPT核对Report":
    st.header("📋 PPT 核对 Report")
    st.write("只检查，不修改 PPT。输出 Price Changed / Missing / Extra 报告。")
    c1, c2 = st.columns(2)
    with c1:
        ppt_file = st.file_uploader("Upload Floorplate PPT / PPTX", type=["pptx", "ppt"], key="audit_ppt")
    with c2:
        project_name = st.text_input("Project Name", value="Floorplate", key="audit_project")
    df, price_df, cfg = excel_controls("audit")

    if ppt_file and price_df is not None and st.button("Generate Audit Report", type="primary"):
        report = audit_only(ppt_file.read(), price_df)
        report_xlsx = report_to_xlsx(report, price_df, project_name=project_name)
        st.success("Report generated.")
        show_metrics(report, price_df)
        st.download_button(
            "Download Audit Report",
            report_xlsx,
            file_name=f"{project_name.replace(' ', '_')}_Audit_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.dataframe(report, use_container_width=True)

elif page == "✅ PPT价格更新":
    st.header("✅ 根据 Excel 自动更新 PPT 里的 Unit 价格")
    st.write("只替换已存在的 `Unit xxx - $xxx` 里的价格；不新增 Missing 户型框；同时生成 Report。")
    st.warning("为了尽量不改变字体和格式，本功能只在同一个 text run 中替换价格。若某些复杂文本无法替换，会在 Report 里显示 Skipped。")
    c1, c2 = st.columns(2)
    with c1:
        ppt_file = st.file_uploader("Upload Floorplate PPT / PPTX", type=["pptx", "ppt"], key="update_ppt")
    with c2:
        project_name = st.text_input("Project Name", value="Floorplate", key="update_project")
    df, price_df, cfg = excel_controls("update")

    if ppt_file and price_df is not None and st.button("Update PPT Prices", type="primary"):
        updated_ppt, report = update_ppt_prices_keep_format(ppt_file.read(), price_df)
        report_xlsx = report_to_xlsx(report, price_df, project_name=project_name)
        st.success("PPT updated and report generated.")
        show_metrics(report, price_df)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Download Updated PPT",
                updated_ppt,
                file_name=f"{project_name.replace(' ', '_')}_Updated_Prices.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        with c2:
            st.download_button(
                "Download Update Report",
                report_xlsx,
                file_name=f"{project_name.replace(' ', '_')}_Price_Update_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.dataframe(report, use_container_width=True)

else:
    st.header(page)
    st.info("This module is planned. Current active modules: PPT核对Report and PPT价格更新.")
