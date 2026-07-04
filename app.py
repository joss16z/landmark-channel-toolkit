import io
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import streamlit as st
from pptx import Presentation

st.set_page_config(page_title="Landmark Channel Toolkit", page_icon="🏢", layout="wide")

# -----------------------------
# Helpers
# -----------------------------

UNIT_RE = re.compile(r"\b(?:Unit|Apt|Apartment)\s*([A-Za-z]{0,3}\.?\s*\d{1,5})\b", re.I)
LOT_UNIT_RE = re.compile(r"\bLot\s*(\d+)\s*[_\- ]+\s*(?:Unit|Apt|Apartment)\s*([A-Za-z]{0,3}\.?\s*\d{1,5})\b", re.I)
PRICE_RE = re.compile(r"\$\s*([0-9][0-9,]*(?:\.\d+)?)")
AREA_RE = re.compile(r"(?<!\d)(\d{1,3})\s*\+\s*(\d{1,3})(?!\d)")


def normalize_unit(value: Any) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    s = re.sub(r"\s+", "", s)
    s = s.replace(".", "")
    # Excel may read unit 201 as 201.0
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s.upper()


def normalize_lot(value: Any) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    m = re.search(r"\d+", s)
    return m.group(0) if m else None


def parse_money(value: Any) -> Optional[int]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return int(round(float(value)))
    s = str(value)
    # Fix malformed PPT prices like $1,015,00 -> $1,015,000 when there are 2 final digits after comma
    m = PRICE_RE.search(s)
    if m:
        raw = m.group(1)
    else:
        raw = s
    raw = raw.replace("$", "").replace("A", "").replace(",", "").strip()
    raw = re.sub(r"[^0-9.]", "", raw)
    if not raw:
        return None
    try:
        return int(round(float(raw)))
    except Exception:
        return None


def parse_ppt_price_from_line(line: str) -> Optional[int]:
    m = PRICE_RE.search(line)
    if not m:
        return None
    token = m.group(1)
    parts = token.split(",")
    # Some PPT labels miss the last zero: 1,015,00 should be 1,015,000
    if len(parts) >= 2 and len(parts[-1]) == 2:
        token = token + "0"
    return parse_money(token)


def parse_number(value: Any) -> Optional[int]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return int(round(float(value)))
    m = re.search(r"\d+", str(value))
    return int(m.group(0)) if m else None


def fmt_money(v: Optional[int]) -> str:
    if v is None:
        return ""
    return f"${v:,.0f}"


def fmt_area(internal: Optional[int], external: Optional[int]) -> str:
    if internal is None and external is None:
        return ""
    return f"{internal or 0} + {external or 0}"


def canon_col(s: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())


def choose_col(columns, candidates: List[str]) -> Optional[str]:
    canon_map = {canon_col(c): c for c in columns}
    candidate_canons = [canon_col(c) for c in candidates]
    for cc in candidate_canons:
        if cc in canon_map:
            return canon_map[cc]
    # contains fallback
    for c in columns:
        ccanon = canon_col(c)
        for cc in candidate_canons:
            if cc and cc in ccanon:
                return c
    return None


# -----------------------------
# Excel parser
# -----------------------------

def read_excel_smart(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"

    # Try normal read first
    try:
        xls = pd.ExcelFile(io.BytesIO(data), engine=engine)
    except Exception:
        xls = pd.ExcelFile(io.BytesIO(data))

    best_df = None
    best_score = -1

    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        raw = raw.dropna(how="all").dropna(axis=1, how="all")
        if raw.empty:
            continue
        max_header_rows = min(20, len(raw))
        for header_idx in range(max_header_rows):
            header_vals = [str(x).strip() if not pd.isna(x) else f"Unnamed_{i}" for i, x in enumerate(raw.iloc[header_idx].tolist())]
            df = raw.iloc[header_idx + 1:].copy()
            df.columns = header_vals[: len(df.columns)]
            cols = list(df.columns)
            score = 0
            if choose_col(cols, ["Apt #", "Unit", "Unit Number", "Apartment", "Apt"]): score += 4
            if choose_col(cols, ["Contract Price", "Price", "List Price", "Gross Price"]): score += 4
            if choose_col(cols, ["Internal", "Internal (sqm)", "Internal Area"]): score += 2
            if choose_col(cols, ["External", "External (sqm)", "External Area"]): score += 2
            if choose_col(cols, ["Lot #", "Lot", "Lot Number"]): score += 1
            if choose_col(cols, ["Status", "Availability"]): score += 1
            if score > best_score:
                best_score = score
                best_df = df

    if best_df is None or best_score < 4:
        raise ValueError("Could not detect a usable table in the Excel file. Please check headers for Unit and Price.")

    return best_df.dropna(how="all")


def parse_excel_available(uploaded_file) -> Tuple[pd.DataFrame, Dict[str, Dict[str, Any]], Dict[str, str]]:
    df = read_excel_smart(uploaded_file)
    cols = list(df.columns)
    colmap = {
        "unit": choose_col(cols, ["Apt #", "Unit", "Unit Number", "Unit No", "Apartment", "Apt"]),
        "lot": choose_col(cols, ["Lot #", "Lot", "Lot Number"]),
        "price": choose_col(cols, ["Contract Price", "Price", "List Price", "Gross Price"]),
        "internal": choose_col(cols, ["Internal", "Internal (sqm)", "Internal Area", "Internal sqm"]),
        "external": choose_col(cols, ["External", "External (sqm)", "External Area", "External sqm"]),
        "status": choose_col(cols, ["Status", "Availability"]),
    }
    if not colmap["unit"] or not colmap["price"]:
        raise ValueError("Could not find Unit and Price columns in Excel.")

    records = []
    for _, row in df.iterrows():
        unit = normalize_unit(row.get(colmap["unit"]))
        if not unit:
            continue
        status = str(row.get(colmap["status"], "Available")).strip() if colmap["status"] else "Available"
        if colmap["status"]:
            if "available" not in status.lower():
                continue
        price = parse_money(row.get(colmap["price"]))
        internal = parse_number(row.get(colmap["internal"])) if colmap["internal"] else None
        external = parse_number(row.get(colmap["external"])) if colmap["external"] else None
        lot = normalize_lot(row.get(colmap["lot"])) if colmap["lot"] else None
        records.append({
            "Unit": unit,
            "Lot": lot,
            "Internal": internal,
            "External": external,
            "Area": fmt_area(internal, external),
            "Excel Price": price,
            "Excel Price Text": fmt_money(price),
            "Status": status,
        })

    out = pd.DataFrame(records).drop_duplicates(subset=["Unit"], keep="first")
    return out, {r["Unit"]: r for r in out.to_dict("records")}, colmap


# -----------------------------
# PPT parser
# -----------------------------

def iter_shapes(shapes):
    for shape in shapes:
        if hasattr(shape, "shapes"):
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def extract_text_from_shape(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(p.text for p in shape.text_frame.paragraphs).strip()


def parse_ppt(uploaded_file) -> Tuple[pd.DataFrame, Dict[str, Dict[str, Any]]]:
    prs = Presentation(io.BytesIO(uploaded_file.getvalue()))
    records: List[Dict[str, Any]] = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in iter_shapes(slide.shapes):
            text = extract_text_from_shape(shape)
            if not text:
                continue
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if not lines:
                continue

            # Case A: grouped text with area header then multiple Lot_Unit - price lines
            current_area = None
            for line in lines:
                area_m = AREA_RE.search(line)
                lot_unit_m = LOT_UNIT_RE.search(line)
                unit_m = UNIT_RE.search(line)
                price = parse_ppt_price_from_line(line)

                if area_m and not unit_m:
                    current_area = (int(area_m.group(1)), int(area_m.group(2)))
                    continue

                if lot_unit_m:
                    unit = normalize_unit(lot_unit_m.group(2))
                    lot = normalize_lot(lot_unit_m.group(1))
                    records.append({
                        "Unit": unit,
                        "Lot": lot,
                        "PPT Price": price,
                        "Internal": current_area[0] if current_area else None,
                        "External": current_area[1] if current_area else None,
                        "Area": fmt_area(current_area[0], current_area[1]) if current_area else "",
                        "Slide": slide_idx,
                        "Raw Text": text,
                    })
                    continue

                if unit_m and price is not None:
                    unit = normalize_unit(unit_m.group(1))
                    area = AREA_RE.search(text)
                    records.append({
                        "Unit": unit,
                        "Lot": None,
                        "PPT Price": price,
                        "Internal": int(area.group(1)) if area else None,
                        "External": int(area.group(2)) if area else None,
                        "Area": fmt_area(int(area.group(1)), int(area.group(2))) if area else "",
                        "Slide": slide_idx,
                        "Raw Text": text,
                    })

            # Case B: Unit on one line, area next, price next
            if any("Unit" in ln for ln in lines):
                for i, line in enumerate(lines):
                    um = UNIT_RE.search(line)
                    if not um:
                        continue
                    unit = normalize_unit(um.group(1))
                    # avoid duplicating if inline price was already captured
                    if parse_ppt_price_from_line(line) is not None:
                        continue
                    window = "\n".join(lines[i:i+5])
                    pm = PRICE_RE.search(window)
                    if not pm:
                        continue
                    price = parse_ppt_price_from_line(window)
                    area = AREA_RE.search(window)
                    # Prevent duplicate same shape/unit/price
                    if not any(r["Unit"] == unit and r["Slide"] == slide_idx and r["Raw Text"] == text for r in records):
                        records.append({
                            "Unit": unit,
                            "Lot": None,
                            "PPT Price": price,
                            "Internal": int(area.group(1)) if area else None,
                            "External": int(area.group(2)) if area else None,
                            "Area": fmt_area(int(area.group(1)), int(area.group(2))) if area else "",
                            "Slide": slide_idx,
                            "Raw Text": text,
                        })

    df = pd.DataFrame(records)
    if df.empty:
        return df, {}
    # Choose first occurrence for comparison; duplicates are reported separately.
    by_unit = df.drop_duplicates(subset=["Unit"], keep="first")
    return df, {r["Unit"]: r for r in by_unit.to_dict("records")}


# -----------------------------
# Audit + report
# -----------------------------

def build_audit(excel_df, excel_map, ppt_df, ppt_map):
    excel_units = set(excel_map.keys())
    ppt_units = set(ppt_map.keys())
    common = sorted(excel_units & ppt_units)

    price_changed = []
    matched = []
    area_changed = []
    lot_changed = []

    for u in common:
        e = excel_map[u]
        p = ppt_map[u]
        e_price, p_price = e.get("Excel Price"), p.get("PPT Price")
        e_int, e_ext = e.get("Internal"), e.get("External")
        p_int, p_ext = p.get("Internal"), p.get("External")
        e_lot, p_lot = e.get("Lot"), p.get("Lot")

        if e_price != p_price:
            price_changed.append({
                "Unit": u,
                "Excel Price": e_price,
                "PPT Price": p_price,
                "Difference": (e_price or 0) - (p_price or 0),
                "Excel Price Text": fmt_money(e_price),
                "PPT Price Text": fmt_money(p_price),
                "PPT Slide": p.get("Slide"),
            })
        else:
            matched.append({"Unit": u, "Price": e_price, "PPT Slide": p.get("Slide")})

        if e_int is not None and e_ext is not None and p_int is not None and p_ext is not None:
            if int(e_int) != int(p_int) or int(e_ext) != int(p_ext):
                area_changed.append({
                    "Unit": u,
                    "Excel Area": fmt_area(e_int, e_ext),
                    "PPT Area": fmt_area(p_int, p_ext),
                    "PPT Slide": p.get("Slide"),
                })

        if e_lot and p_lot and str(e_lot) != str(p_lot):
            lot_changed.append({
                "Unit": u,
                "Excel Lot": e_lot,
                "PPT Lot": p_lot,
                "PPT Slide": p.get("Slide"),
            })

    missing = []
    for u in sorted(excel_units - ppt_units):
        e = excel_map[u]
        missing.append({
            "Unit": u,
            "Lot": e.get("Lot"),
            "Area": e.get("Area"),
            "Excel Price": e.get("Excel Price"),
            "Excel Price Text": fmt_money(e.get("Excel Price")),
            "Status": e.get("Status"),
        })

    extra = []
    for u in sorted(ppt_units - excel_units):
        p = ppt_map[u]
        extra.append({
            "Unit": u,
            "Lot": p.get("Lot"),
            "Area": p.get("Area"),
            "PPT Price": p.get("PPT Price"),
            "PPT Price Text": fmt_money(p.get("PPT Price")),
            "PPT Slide": p.get("Slide"),
        })

    duplicates = pd.DataFrame()
    if not ppt_df.empty:
        dup_counts = ppt_df.groupby("Unit").size().reset_index(name="Count")
        dup_units = set(dup_counts.loc[dup_counts["Count"] > 1, "Unit"])
        duplicates = ppt_df[ppt_df["Unit"].isin(dup_units)].copy() if dup_units else pd.DataFrame()

    summary = pd.DataFrame([
        {"Metric": "Excel Available", "Count": len(excel_units)},
        {"Metric": "PPT Tagged", "Count": len(ppt_units)},
        {"Metric": "Matched Price OK", "Count": len(matched)},
        {"Metric": "Price Changed", "Count": len(price_changed)},
        {"Metric": "Missing in PPT", "Count": len(missing)},
        {"Metric": "Extra in PPT", "Count": len(extra)},
        {"Metric": "Area Changed", "Count": len(area_changed)},
        {"Metric": "Lot Changed", "Count": len(lot_changed)},
        {"Metric": "Duplicate PPT Labels", "Count": 0 if duplicates.empty else duplicates["Unit"].nunique()},
    ])

    return {
        "Summary": summary,
        "Price Changed": pd.DataFrame(price_changed),
        "Missing in PPT": pd.DataFrame(missing),
        "Extra in PPT": pd.DataFrame(extra),
        "Area Changed": pd.DataFrame(area_changed),
        "Lot Changed": pd.DataFrame(lot_changed),
        "Matched": pd.DataFrame(matched),
        "Duplicates in PPT": duplicates,
        "Raw PPT Labels": ppt_df,
        "Raw Excel Available": excel_df,
    }


def make_report_xlsx(audit: Dict[str, pd.DataFrame], project_name: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
        money_fmt = workbook.add_format({"num_format": "$#,##0", "border": 1})
        normal_fmt = workbook.add_format({"border": 1})
        red_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1})
        green_fmt = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1})
        title_fmt = workbook.add_format({"bold": True, "font_size": 16})

        for sheet_name, df in audit.items():
            if df is None or df.empty:
                df = pd.DataFrame(columns=["No issues found"])
            # Excel sheet names max 31 chars
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False, startrow=2 if sheet_name == "Summary" else 0)
            ws = writer.sheets[safe_name]

            if sheet_name == "Summary":
                ws.write(0, 0, f"{project_name} - Floorplate Update Report", title_fmt)
                ws.write(1, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
                header_row = 2
            else:
                header_row = 0

            # headers
            for col_num, value in enumerate(df.columns):
                ws.write(header_row, col_num, value, header_fmt)

            # formatting
            for idx, col in enumerate(df.columns):
                col_width = max(12, min(45, max([len(str(col))] + [len(str(x)) for x in df[col].head(200).fillna("").tolist()])))
                ws.set_column(idx, idx, col_width)
                if "Price" in str(col) or "Difference" in str(col):
                    ws.set_column(idx, idx, 16, money_fmt)

            if sheet_name in {"Price Changed", "Missing in PPT", "Extra in PPT", "Area Changed", "Lot Changed"}:
                ws.conditional_format(1, 0, max(1, len(df)), max(0, len(df.columns)-1), {"type": "no_blanks", "format": red_fmt})
            if sheet_name == "Summary" and len(df) > 0:
                # Highlight zero problem counts green, non-zero red for issue rows
                for r in range(3, 3 + len(df)):
                    metric = df.iloc[r-3, 0]
                    count = df.iloc[r-3, 1]
                    if metric in {"Price Changed", "Missing in PPT", "Extra in PPT", "Area Changed", "Lot Changed", "Duplicate PPT Labels"}:
                        ws.write(r, 1, count, green_fmt if count == 0 else red_fmt)

            ws.freeze_panes(header_row + 1, 0)

    output.seek(0)
    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------

st.title("🏢 Landmark Channel Toolkit")
st.caption("V1 — PPT核对Python / Floorplate Update Report")

with st.sidebar:
    st.header("功能")
    st.markdown("✅ **PPT核对Python**")
    st.markdown("🚧 Floorplate Generator")
    st.markdown("🚧 Price Compare")
    st.markdown("🚧 Agent Package Generator")

st.subheader("📑 PPT核对Python")
st.write("上传 Floorplate PPT 和 Price List Excel，自动检查价格、Missing、Extra、面积和 Lot。")

col1, col2 = st.columns(2)
with col1:
    ppt_file = st.file_uploader("Upload Floorplate PPT / PPTX", type=["ppt", "pptx"])
with col2:
    excel_file = st.file_uploader("Upload Price List Excel", type=["xls", "xlsx"])

project_name = st.text_input("Project Name", value="Floorplate")

if st.button("Generate Floorplate Update Report", type="primary"):
    if not ppt_file or not excel_file:
        st.error("Please upload both PPT and Excel files.")
    else:
        try:
            with st.spinner("Reading Excel..."):
                excel_df, excel_map, colmap = parse_excel_available(excel_file)
            with st.spinner("Reading PPT..."):
                ppt_df, ppt_map = parse_ppt(ppt_file)
            with st.spinner("Building report..."):
                audit = build_audit(excel_df, excel_map, ppt_df, ppt_map)
                report_bytes = make_report_xlsx(audit, project_name)

            summary_df = audit["Summary"]
            st.success("Report generated successfully.")

            c1, c2, c3, c4, c5 = st.columns(5)
            summary_dict = dict(zip(summary_df["Metric"], summary_df["Count"]))
            c1.metric("Excel Available", summary_dict.get("Excel Available", 0))
            c2.metric("PPT Tagged", summary_dict.get("PPT Tagged", 0))
            c3.metric("Price Changed", summary_dict.get("Price Changed", 0))
            c4.metric("Missing", summary_dict.get("Missing in PPT", 0))
            c5.metric("Extra", summary_dict.get("Extra in PPT", 0))

            st.markdown("### Summary")
            st.dataframe(summary_df, use_container_width=True)

            with st.expander("Excel column mapping detected"):
                st.json({k: str(v) for k, v in colmap.items()})

            tabs = st.tabs(["Price Changed", "Missing", "Extra", "Area Changed", "Lot Changed"])
            for tab, key in zip(tabs, ["Price Changed", "Missing in PPT", "Extra in PPT", "Area Changed", "Lot Changed"]):
                with tab:
                    df = audit[key]
                    if df.empty:
                        st.success("No issues found.")
                    else:
                        st.dataframe(df, use_container_width=True)

            filename_project = re.sub(r"[^A-Za-z0-9]+", "_", project_name).strip("_") or "Floorplate"
            today = datetime.now().strftime("%Y%m%d")
            st.download_button(
                label="⬇️ Download Floorplate Update Report",
                data=report_bytes,
                file_name=f"{filename_project}_Floorplate_Update_Report_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("Failed to generate report.")
            st.exception(e)
