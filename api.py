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

                # THE BUG FIX: using sheet.lower() instead of sheet.upper()
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
        total_sites = df_filtered["Location"].nunique()
        avg_alarms = round(total_alarms / unique_nodes, 1) if unique_nodes > 0 else 0
        total_crit = int((df_filtered["Severity_Label"] == "Critical").sum())

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
            {"name": "Major", "value": vc.get("Major", 0)},
            {"name": "Minor", "value": vc.get("Minor", 0)},
        ]

        # ── Per-location donut data ───────────────────────────────────────────
        donut_data = []
        for (loc, net), grp in df_filtered.groupby(["Location", "network_type"]):
            v = grp["Severity_Label"].value_counts().to_dict()
            donut_data.append({
                "location": loc,
                "network_type": net,
                "Critical": v.get("Critical", 0),
                "Major": v.get("Major", 0),
                "Minor": v.get("Minor", 0),
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
                "total_alarms": total_alarms,
                "avg_alarms": avg_alarms,
                "total_sites": total_sites,
                "critical_alarms": total_crit,
                "unique_nodes": int(unique_nodes),
            },
            "pie_chart_data": pie_chart_data,
            "chart_data": chart_data,
            "donut_data": donut_data,
            "all_locations": all_locations,
            "raw_alarms": raw_alarms,
            "investigator_logs": investigator_logs,
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"status": "error", "message": str(e)}