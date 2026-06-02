from fastapi import FastAPI, UploadFile, File, Query
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io
import re

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Location decoder ──────────────────────────────────────────────────────────
DECODER = {
    "AM": "Ashok Marg",
    "GN": "Gomti Nagar",
    "GK": "Gomti Nagar",
    "CS": "Gomti Nagar",
    "TI": "Gomti Nagar",
    "GW": "Gomti Nagar",
    "GT": "Gomti Nagar",
    "VR": "Varanasi",
}


def decode_location(node_id: str) -> str:
    s = str(node_id).strip().upper()

    m = re.search(r'([A-Z]{2})\d+', s)
    if m and m.group(1) in DECODER:
        return DECODER[m.group(1)]

    for code, location in DECODER.items():
        if code in s:
            return location

    clean = re.sub(r'[^A-Z]', '', s)[:12]
    for i in range(len(clean) - 1):
        pair = clean[i:i + 2]
        if pair in DECODER:
            return DECODER[pair]

    return s


# ─────────────────────────────────────────────────────────────────────────────
# Helper: read UPE Node Count from the SUMMARY sheet
# Layout (from screenshot):
#   Row 2  → headers:  CIRCLE | ACTIVE_ALARM | ALARM<24 | ALARM>24 | ALARM>48 | Node Count | ...
#   Row 20 → data:     UPE    | 157          | 44       | 20       | 93       | 87         | ...
#
# Strategy:
#   1. Read the SUMMARY sheet with no header (header=None) so row indices are stable.
#   2. Find which row has "UPE" in column A (index 0).
#   3. Find which column has "Node Count" in row 2 (index 1, the header row).
#   4. Return the integer at [UPE row, Node Count col].
#   Falls back to None if anything is missing or fails.
# ─────────────────────────────────────────────────────────────────────────────
def get_upe_node_count(xls: pd.ExcelFile) -> int | None:
    # Step 1: find the SUMMARY sheet (case-insensitive)
    summary_sheet = next(
        (s for s in xls.sheet_names if s.strip().upper() == "SUMMARY"),
        None
    )
    if summary_sheet is None:
        return None

    try:
        df = pd.read_excel(xls, sheet_name=summary_sheet, header=None)

        # Step 2: find the row where column A (index 0) == "UPE"
        upe_row_idx = None
        for i, val in enumerate(df.iloc[:, 0]):
            if str(val).strip().upper() == "UPE":
                upe_row_idx = i
                break
        if upe_row_idx is None:
            return None

        # Step 3: find the column whose header (row index 1) == "Node Count"
        # The header row is row index 1 based on the screenshot (row 2 in Excel = index 1)
        node_count_col_idx = None
        for col_idx in range(len(df.columns)):
            header_val = str(df.iloc[1, col_idx]).strip()
            if "node count" in header_val.lower() or header_val.lower() == "node count":
                node_count_col_idx = col_idx
                break

        # Fallback: if header search fails, use column F (index 5) which is
        # the Node Count column as seen in the screenshot
        if node_count_col_idx is None:
            node_count_col_idx = 5

        # Step 4: read and return the value
        raw_val = df.iloc[upe_row_idx, node_count_col_idx]
        node_count = int(float(str(raw_val).replace(",", "").strip()))
        if node_count > 0:
            return node_count

    except Exception as e:
        print(f"[get_upe_node_count] {e}")

    return None


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — List available sheets
# ─────────────────────────────────────────────────────────────────────────────
@app.post("/api/list-sheets")
async def list_sheets(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        if file.filename.endswith('.csv'):
            return {"status": "success", "sheets": ["CSV Data"], "is_csv": True}
        engine = 'pyxlsb' if file.filename.endswith('.xlsb') else 'openpyxl'
        xls = pd.ExcelFile(io.BytesIO(contents), engine=engine)
        return {"status": "success", "sheets": xls.sheet_names, "is_csv": False}
    except Exception as e:
        return {"status": "error", "message": str(e)}


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — Full processing with selected sheets + optional location filter
# ─────────────────────────────────────────────────────────────────────────────
@app.post("/api/process-alarms")
async def process_alarms(
        file: UploadFile = File(...),
        selected_sheets: str = Query(default=""),
        selected_locations: str = Query(default=""),
        investigator_loc: str = Query(default=""),
        investigator_sev: str = Query(default="All"),
        investigator_net: str = Query(default="All"),
):
    try:
        contents = await file.read()
        sheet_list = [s.strip() for s in selected_sheets.split(",") if s.strip()]
        loc_filter = [l.strip() for l in selected_locations.split(",") if l.strip()]

        # ── Load ──────────────────────────────────────────────────────────────
        if file.filename.endswith('.csv'):
            df_raw = pd.read_csv(io.BytesIO(contents))
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.lower()
            df_raw["network_type"] = "Unknown"
            xls = None
        else:
            engine = 'pyxlsb' if file.filename.endswith('.xlsb') else 'openpyxl'
            xls = pd.ExcelFile(io.BytesIO(contents), engine=engine)

            if not sheet_list:
                sheet_list = [s for s in xls.sheet_names
                              if "mss" in s.lower() or "mgw" in s.lower()]
            if not sheet_list:
                return {"status": "error", "message": "No sheets selected or found."}

            frames = []
            for sheet in sheet_list:
                df_t = pd.read_excel(xls, sheet_name=sheet)
                df_t.columns = df_t.columns.astype(str).str.strip().str.lower()

                if "mss" in sheet.lower():
                    df_t["network_type"] = "MSS"
                elif "mgw" in sheet.lower():
                    df_t["network_type"] = "MGW"
                else:
                    df_t["network_type"] = sheet
                frames.append(df_t)

            if not frames:
                return {"status": "error", "message": "Could not load data from the selected sheets."}
            df_raw = pd.concat(frames, ignore_index=True)

        # ── Validate ──────────────────────────────────────────────────────────
        if "nodename" not in df_raw.columns or "severity" not in df_raw.columns:
            return {"status": "error",
                    "message": "Could not find 'NodeName' or 'Severity' columns."}

        # ── Clean & decode ────────────────────────────────────────────────────
        df_raw["Location"] = df_raw["nodename"].apply(decode_location)
        df_raw["severity"] = pd.to_numeric(df_raw["severity"], errors="coerce")
        df_raw["Severity_Label"] = df_raw["severity"].map({1: "Critical", 2: "Major", 3: "Minor"})
        df_clean = df_raw.dropna(subset=["Location", "Severity_Label"]).copy()

        # ── All locations ─────────────────────────────────────────────────────
        all_locations = sorted(df_clean["Location"].unique().tolist())

        if loc_filter:
            df_filtered = df_clean[df_clean["Location"].isin(loc_filter)].copy()
        else:
            df_filtered = df_clean.copy()

        # ── KPIs ──────────────────────────────────────────────────────────────
        total_alarms = len(df_filtered)
        unique_nodes = df_filtered["nodename"].nunique()
        total_sites  = df_filtered["Location"].nunique()
        total_crit   = int((df_filtered["Severity_Label"] == "Critical").sum())

        # ── Official node count from SUMMARY sheet → UPE row → Node Count col ─
        # avg/node = total_alarms ÷ official_node_count
        # Falls back to unique_nodes (counted from alarm rows) if not found.
        official_node_count = get_upe_node_count(xls) if xls is not None else None
        denominator = official_node_count if official_node_count else int(unique_nodes)
        avg_alarms  = round(total_alarms / denominator, 1) if denominator > 0 else 0

        # ── Grouped chart data ────────────────────────────────────────────────
        df_grouped = (
            df_filtered
            .groupby(["Location", "network_type", "Severity_Label"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        for col in ["Critical", "Major", "Minor"]:
            if col not in df_grouped.columns:
                df_grouped[col] = 0
        df_grouped["Grand Total"] = df_grouped["Critical"] + df_grouped["Major"] + df_grouped["Minor"]
        chart_data = df_grouped.to_dict(orient="records")

        # ── Pie / overall breakdown ───────────────────────────────────────────
        vc = df_filtered["Severity_Label"].value_counts().to_dict()
        pie_chart_data = [
            {"name": "Critical", "value": vc.get("Critical", 0)},
            {"name": "Major",    "value": vc.get("Major",    0)},
            {"name": "Minor",    "value": vc.get("Minor",    0)},
        ]

        # ── Per-location donut data ───────────────────────────────────────────
        donut_data = []
        for (loc, net), grp in df_filtered.groupby(["Location", "network_type"]):
            v = grp["Severity_Label"].value_counts().to_dict()
            donut_data.append({
                "location":     loc,
                "network_type": net,
                "Critical":     v.get("Critical", 0),
                "Major":        v.get("Major",    0),
                "Minor":        v.get("Minor",    0),
            })

        # ── Alarm text column ─────────────────────────────────────────────────
        text_cols = [c for c in df_filtered.columns
                     if any(k in c for k in ("problem", "text", "description"))]
        alarm_text_col = text_cols[0] if text_cols else "nodename"

        # ── Recent raw alarms ─────────────────────────────────────────────────
        raw_alarms = (
            df_filtered.head(150)
            .fillna("N/A")
            .rename(columns={alarm_text_col: "alarm_text"})
            .to_dict(orient="records")
        )

        # ── Investigator drill-down ───────────────────────────────────────────
        investigator_logs = []
        if investigator_loc:
            inv = df_clean[df_clean["Location"] == investigator_loc].copy()
            if investigator_sev != "All":
                inv = inv[inv["Severity_Label"] == investigator_sev]
            if investigator_net != "All":
                inv = inv[inv["network_type"] == investigator_net]

            keep = ["nodename", "network_type", "Severity_Label"]
            if alarm_text_col != "nodename":
                keep.append(alarm_text_col)
            if "aging in hrs" in df_clean.columns:
                keep.append("aging in hrs")
            if "aging in min" in df_clean.columns:
                keep.append("aging in min")

            investigator_logs = (
                inv[keep]
                .fillna("N/A")
                .rename(columns={alarm_text_col: "alarm_text"})
                .to_dict(orient="records")
            )

        # ── Response ──────────────────────────────────────────────────────────
        return {
            "status": "success",
            "kpis": {
                "total_alarms":        total_alarms,
                "avg_alarms":          float(avg_alarms),
                "total_sites":         total_sites,
                "critical_alarms":     total_crit,
                "unique_nodes":        int(unique_nodes),
                "official_node_count": official_node_count,  # 87 from SUMMARY sheet, or null
            },
            "pie_chart_data":    pie_chart_data,
            "chart_data":        chart_data,
            "donut_data":        donut_data,
            "all_locations":     all_locations,
            "raw_alarms":        raw_alarms,
            "investigator_logs": investigator_logs,
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"status": "error", "message": str(e)}