import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re

# 1. Page Setup & Premium CSS Injection
st.set_page_config(page_title="Vi Pan-India Dashboard", layout="wide", page_icon="📡")

st.markdown(
    """
    <style>
    /* Hide Streamlit clutters */
    [data-testid="stStatusWidget"] {visibility: hidden;}
    [data-testid="stHeader"] {display: none;}
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }

    /* PREMIUM KPI CARDS */
    [data-testid="metric-container"] {
        background-color: #1E1E24; border: 1px solid #333333; padding: 15px;
        border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        border-left: 4px solid #FF4B4B; transition: transform 0.2s ease-in-out;
    }
    [data-testid="metric-container"]:hover { transform: translateY(-2px); }
    [data-testid="stMetricLabel"] { color: #A0A0A0; font-size: 1.05rem; font-weight: 500; padding-bottom: 5px; }
    [data-testid="stMetricValue"] { color: #FFFFFF; font-size: 2.2rem; font-weight: 700; }
    [data-testid="stSidebar"] { background-color: #121216; border-right: 1px solid #2A2A35; }
    h1, h2, h3 { color: #E0E0E0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }

    /* --- THE NEW NETWORK RADAR ANIMATION --- */
    .network-radar {
        position: relative;
        width: 120px;
        height: 120px;
        margin: 0 auto 30px auto;
    }
    .network-radar::before, .network-radar::after {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; bottom: 0;
        border-radius: 50%;
        border: 2px solid #FF4B4B; /* Vi Red rings */
        animation: pulse-ring 2s linear infinite;
        opacity: 0;
    }
    .network-radar::after {
        animation-delay: 1s;
    }
    .core-node {
        position: absolute;
        top: 50%; left: 50%;
        transform: translate(-50%, -50%);
        width: 24px;
        height: 24px;
        background-color: #FFA500; /* Vi Yellow core */
        border-radius: 50%;
        box-shadow: 0 0 20px #FFA500;
    }
    @keyframes pulse-ring {
        0% { transform: scale(0.3); opacity: 1; border-width: 4px; }
        100% { transform: scale(2.5); opacity: 0; border-width: 1px; }
    }
    .splash-title {
        color: #FAFAFA;
        font-size: 3rem;
        font-weight: 800;
        letter-spacing: 1px;
        margin: 0;
    }
    .fade-text {
        color: #A0A0A0; 
        font-size: 1.1rem; 
        letter-spacing: 3px; 
        margin-top: 10px;
        animation: fadein 2s infinite alternate;
    }
    @keyframes fadein {
        0% { opacity: 0.3; }
        100% { opacity: 1; }
    }
    </style>
    """,
    unsafe_allow_html=True
)

# 2. File Uploader
uploaded_file = st.file_uploader("📂 Upload the Raw Alarm Log (.xlsb, .xlsx, .csv)", type=["xlsb", "xlsx", "csv"])

# --- THE SPLASH SCREEN ENGINE ---
splash = st.empty()

if uploaded_file is None:
    with splash.container():
        st.markdown(
            """
            <div style='display: flex; flex-direction: column; justify-content: center; align-items: center; height: 60vh;'>
                <div class='network-radar'>
                    <div class='core-node'></div>
                </div>
                <h1 class='splash-title'><span style='color: #FF4B4B;'>Vi</span> Network Automation</h1>
                <p class='fade-text'>AWAITING RAW ALARM DATA...</p>
            </div>
            """,
            unsafe_allow_html=True
        )
else:
    splash.empty()
# --- END SPLASH SCREEN ---

# Main Title (Shows up after upload)
if uploaded_file is not None:
    st.markdown("<h1 style='text-align: center; color: white;'>📡 Vi Alarm Dashboard</h1>",
                unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #A0A0A0;'>Live Enterprise Alarm Monitoring & Analytics</p>",
                unsafe_allow_html=True)
    st.divider()

# --- INITIALIZE SYNCED CLICK STATES ---
if "clicked_loc" not in st.session_state: st.session_state.clicked_loc = None
if "clicked_sev" not in st.session_state: st.session_state.clicked_sev = "All"
if "clicked_net" not in st.session_state: st.session_state.clicked_net = "All"

# 4. The Backend Engine
if uploaded_file is not None:

    if uploaded_file.name.endswith('.csv'):
        df_raw = pd.read_csv(uploaded_file)
        df_raw["network_type"] = "Unknown"
    else:
        engine = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else 'openpyxl'
        xls = pd.ExcelFile(uploaded_file, engine=engine)
        available_sheets = xls.sheet_names

        st.subheader("📄 Select Sheets to Process")
        selected_sheets = st.multiselect(
            "Which tabs contain the raw alarms? (e.g., UPE_MSS, UPE_MGW)",
            options=available_sheets,
            default=[]
        )

        if not selected_sheets:
            st.info("👆 Please select your network sheets from the dropdown above to continue.")
            st.stop()

        all_dataframes = []
        for sheet in selected_sheets:
            df_temp = pd.read_excel(xls, sheet_name=sheet)
            if "MSS" in sheet.upper():
                df_temp["network_type"] = "MSS"
            elif "MGW" in sheet.upper():
                df_temp["network_type"] = "MGW"
            else:
                df_temp["network_type"] = sheet
            all_dataframes.append(df_temp)

        df_raw = pd.concat(all_dataframes, ignore_index=True)

    # --- CLEANING & TRANSLATOR ---
    df_raw.columns = df_raw.columns.astype(str).str.strip().str.lower()

    if "nodename" not in df_raw.columns or "severity" not in df_raw.columns:
        st.error("❌ Could not find 'NodeName' or 'Severity' columns! Are you sure this is the raw alarm dump?")
        st.stop()

    DECODER = {"AM": "Ashok Marg", "GN": "Gomti Nagar", "GK": "Gomti Nagar", "GT": "Gomti Nagar", "VR": "Varanasi"}


    def decode_location(node_id):
        match = re.search(r'([A-Za-z]{2})\d+$', str(node_id).strip())
        if match: return DECODER.get(match.group(1).upper(), "Gomti Nagar")
        return str(node_id).strip()


    df_raw["Location"] = df_raw["nodename"].apply(decode_location)
    sev_mapping = {1: "Critical", 2: "Major", 3: "Minor"}
    df_raw["severity"] = pd.to_numeric(df_raw["severity"], errors='coerce')
    df_raw["Severity_Label"] = df_raw["severity"].map(sev_mapping)

    df_clean = df_raw.dropna(subset=["Location", "Severity_Label"]).copy()

    df_grouped = df_clean.groupby(["Location", "network_type", "Severity_Label"]).size().unstack(
        fill_value=0).reset_index()
    for col in ["Critical", "Major", "Minor"]:
        if col not in df_grouped.columns: df_grouped[col] = 0
    df_grouped["Grand Total"] = df_grouped["Critical"] + df_grouped["Major"] + df_grouped["Minor"]

    all_locations = sorted(df_grouped["Location"].unique().tolist())

    # --- THE SIDEBAR INVESTIGATOR ---
    if st.session_state.clicked_loc is not None:
        st.sidebar.title("🔍 Investigator")
        st.sidebar.caption("Drill down into specific hardware logs.")

        if st.sidebar.button("❌ Close Panel", use_container_width=True, type="primary"):
            st.session_state.clicked_loc = None
            st.session_state.clicked_sev = "All"
            st.session_state.clicked_net = "All"
            st.rerun()


        def sync_loc_from_sidebar():
            st.session_state.clicked_loc = st.session_state.sb_loc


        def sync_sev_from_sidebar():
            st.session_state.clicked_sev = st.session_state.sb_sev


        def sync_net_from_sidebar():
            st.session_state.clicked_net = st.session_state.sb_net


        if st.session_state.clicked_loc in all_locations: st.session_state.sb_loc = st.session_state.clicked_loc
        if st.session_state.clicked_sev in ["All", "Critical", "Major",
                                            "Minor"]: st.session_state.sb_sev = st.session_state.clicked_sev
        if st.session_state.clicked_net in ["All", "MSS", "MGW"]: st.session_state.sb_net = st.session_state.clicked_net

        investigator_loc = st.sidebar.selectbox("📍 Target Location:", all_locations, key="sb_loc",
                                                on_change=sync_loc_from_sidebar)
        investigator_net = st.sidebar.selectbox("📶 Network Type:", ["All", "MSS", "MGW"], key="sb_net",
                                                on_change=sync_net_from_sidebar)
        investigator_sev = st.sidebar.selectbox("⚠️ Severity Filter:", ["All", "Critical", "Major", "Minor"],
                                                key="sb_sev", on_change=sync_sev_from_sidebar)

        raw_display = df_clean[df_clean["Location"] == investigator_loc]
        if investigator_sev != "All":
            raw_display = raw_display[raw_display["Severity_Label"] == investigator_sev]
        if investigator_net != "All":
            raw_display = raw_display[raw_display["network_type"] == investigator_net]

        display_map = {"nodename": "Node ID", "network_type": "Net", "Severity_Label": "Severity",
                       "alarm_text": "Alarm Text"}
        if "aging in hrs" in df_clean.columns: display_map["aging in hrs"] = "Hrs"
        if "aging in min" in df_clean.columns: display_map["aging in min"] = "Mins"

        final_table = raw_display[list(display_map.keys())].rename(columns=display_map)
        st.sidebar.markdown(f"**Live Logs:** {investigator_loc}")
        st.sidebar.dataframe(final_table, use_container_width=True, hide_index=True, height=600)

    # --- TOP CONTROLS & KPIS ---
    st.subheader("🌍 Executive Summary")
    selected_locations = st.multiselect("Active Network Nodes:", options=all_locations,
                                        default=all_locations[:4] if len(all_locations) >= 4 else all_locations)

    filtered_df = df_grouped[df_grouped["Location"].isin(selected_locations)]
    raw_filtered_df = df_clean[df_clean["Location"].isin(selected_locations)]

    df_mss = filtered_df[filtered_df["network_type"] == "MSS"]
    df_mgw = filtered_df[filtered_df["network_type"] == "MGW"]

    total_alarms = len(raw_filtered_df)
    unique_nodes = raw_filtered_df["nodename"].nunique()
    total_sites = raw_filtered_df["Location"].nunique()

    avg_alarms_per_node = round(total_alarms / unique_nodes, 1) if unique_nodes > 0 else 0

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric(label="Total Active Alarms", value=f"{total_alarms:,}")
    kpi2.metric(label="Avg Alarms per Node", value=f"{avg_alarms_per_node}")
    kpi3.metric(label="Monitored Sites", value=total_sites)

    st.divider()

    # --- INTERACTIVE SPLIT SEVERITY CHARTS ---
    st.subheader("🔍 Severity Analysis")
    st.caption("Select a severity filter below, or click any bar on the chart to instantly fetch raw hardware logs.")

    selected_chart_sev = st.selectbox("Global Bar Chart Filter:", ["All", "Critical", "Major", "Minor"])

    col_bar1, col_bar2 = st.columns(2)


    def draw_interactive_bar(df_subset, title, col, network_tag):
        if df_subset.empty:
            col.info(f"No data available for {title}.")
            return None

        bar_df = df_subset.copy()
        melted_bar_df = bar_df.melt(id_vars=["Location"], value_vars=["Critical", "Major", "Minor"],
                                    var_name="Severity", value_name="Count")
        melted_bar_df = melted_bar_df[melted_bar_df["Count"] > 0]

        if selected_chart_sev != "All":
            melted_bar_df = melted_bar_df[melted_bar_df["Severity"] == selected_chart_sev]

        if not melted_bar_df.empty:
            fig_bar = px.bar(
                melted_bar_df, x="Location", y="Count", color="Severity", barmode="group", text_auto=True,
                custom_data=["Severity"],
                title=f"{title}",
                color_discrete_map={"Critical": "#FF4B4B", "Major": "#FFA500", "Minor": "#00CC96"}
            )

            fig_bar.update_layout(
                xaxis_title="",
                yaxis_title="Alarm Count",
                clickmode="event+select",
                height=400,
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=20, r=20, t=40, b=20)
            )

            if st.session_state.clicked_loc and st.session_state.clicked_net in ["All", network_tag]:
                for trace in fig_bar.data:
                    if st.session_state.clicked_sev == "All" or trace.name == st.session_state.clicked_sev:
                        selected_indices = [i for i, loc in enumerate(trace.x) if loc == st.session_state.clicked_loc]
                        trace.selectedpoints = selected_indices if selected_indices else []
                    else:
                        trace.selectedpoints = []

            return col.plotly_chart(fig_bar, use_container_width=True, on_select="rerun", selection_mode="points")
        return None


    sel_mss = draw_interactive_bar(df_mss, "🖥️ MSS Network", col_bar1, "MSS")
    sel_mgw = draw_interactive_bar(df_mgw, "📡 MGW Network", col_bar2, "MGW")

    chart_was_clicked = False
    if sel_mss and len(sel_mss.selection["points"]) > 0:
        new_loc = sel_mss.selection["points"][0]["x"]
        new_sev = sel_mss.selection["points"][0]["customdata"][0]
        if st.session_state.clicked_loc != new_loc or st.session_state.clicked_sev != new_sev or st.session_state.clicked_net != "MSS":
            st.session_state.clicked_loc = new_loc
            st.session_state.clicked_sev = new_sev
            st.session_state.clicked_net = "MSS"
            chart_was_clicked = True
    elif sel_mgw and len(sel_mgw.selection["points"]) > 0:
        new_loc = sel_mgw.selection["points"][0]["x"]
        new_sev = sel_mgw.selection["points"][0]["customdata"][0]
        if st.session_state.clicked_loc != new_loc or st.session_state.clicked_sev != new_sev or st.session_state.clicked_net != "MGW":
            st.session_state.clicked_loc = new_loc
            st.session_state.clicked_sev = new_sev
            st.session_state.clicked_net = "MGW"
            chart_was_clicked = True
    if chart_was_clicked:
        st.rerun()

    st.divider()


    # --- TABLES ---
    def append_totals_and_style(df):
        if df.empty: return df
        display_df = df[["Location", "Critical", "Major", "Minor", "Grand Total"]].copy()

        totals = display_df[["Critical", "Major", "Minor", "Grand Total"]].sum()
        total_row = pd.DataFrame(
            [["TOTAL", totals["Critical"], totals["Major"], totals["Minor"], totals["Grand Total"]]],
            columns=["Location", "Critical", "Major", "Minor", "Grand Total"])
        display_df = pd.concat([display_df, total_row], ignore_index=True)

        sums = {"Critical": totals["Critical"], "Major": totals["Major"], "Minor": totals["Minor"]}
        ranked_severities = sorted(sums.keys(), key=lambda k: sums[k], reverse=True)

        data_only = display_df[display_df["Location"] != "TOTAL"]
        max_val = data_only["Grand Total"].max() if not data_only.empty else -1
        min_val = data_only["Grand Total"].min() if not data_only.empty else -1

        def row_styler(row):
            styles = [''] * len(row)
            is_total_row = (row["Location"] == "TOTAL")

            for i, col in enumerate(display_df.columns):
                if is_total_row:
                    if col == "Location":
                        styles[
                            i] = 'background-color: #121216; color: white; font-weight: bold; border-top: 2px solid #555;'
                    elif col == ranked_severities[0]:
                        styles[
                            i] = 'background-color: #FF4B4B; color: black; font-weight: bold; border-top: 2px solid #555;'
                    elif col == ranked_severities[1]:
                        styles[
                            i] = 'background-color: #FFA500; color: black; font-weight: bold; border-top: 2px solid #555;'
                    elif col == ranked_severities[2]:
                        styles[
                            i] = 'background-color: #00CC96; color: black; font-weight: bold; border-top: 2px solid #555;'
                    elif col == "Grand Total":
                        styles[
                            i] = 'background-color: #1E88E5; color: white; font-weight: bold; font-size: 1.1em; border-top: 2px solid #555;'
                else:
                    if col == "Grand Total":
                        if row[col] == max_val:
                            styles[i] = 'background-color: rgba(255, 75, 75, 0.8); color: white; font-weight: bold'
                        elif row[col] == min_val:
                            styles[i] = 'background-color: rgba(0, 204, 150, 0.8); color: white; font-weight: bold'
                        else:
                            styles[i] = 'background-color: rgba(255, 165, 0, 0.8); color: white; font-weight: bold'
            return styles

        return display_df.style.apply(row_styler, axis=1)


    col_table1, col_table2 = st.columns(2)
    with col_table1:
        st.subheader("🖥️ MSS Alarm Directory")
        if not df_mss.empty: st.dataframe(append_totals_and_style(df_mss), use_container_width=True, hide_index=True)
    with col_table2:
        st.subheader("📡 MGW Alarm Directory")
        if not df_mgw.empty: st.dataframe(append_totals_and_style(df_mgw), use_container_width=True, hide_index=True)

    st.divider()


    # --- UNIFIED DONUTS ---
    def draw_unified_donuts(df_subset, title):
        st.subheader(title)
        if df_subset.empty: return
        locations = df_subset['Location'].tolist()
        num_charts = len(locations)
        if num_charts == 0: return

        specs = [[{"type": "domain"}] * num_charts]
        fig = make_subplots(rows=1, cols=num_charts, subplot_titles=locations, specs=specs)
        color_map = {"Critical": "#FF4B4B", "Major": "#FFA500", "Minor": "#00CC96"}

        for i, row in df_subset.reset_index(drop=True).iterrows():
            chart_data = pd.DataFrame(
                {"Severity": ["Critical", "Major", "Minor"], "Count": [row["Critical"], row["Major"], row["Minor"]]})
            chart_data = chart_data[chart_data["Count"] > 0]

            if not chart_data.empty:
                pull_array = [0.1 if s == "Critical" else 0 for s in chart_data["Severity"]]
                colors = [color_map[s] for s in chart_data["Severity"]]
                fig.add_trace(
                    go.Pie(labels=chart_data["Severity"], values=chart_data["Count"], hole=0.45, pull=pull_array,
                           marker_colors=colors, textinfo='label+value', name=locations[i]),
                    row=1, col=i + 1
                )

        fig.update_layout(
            showlegend=False, height=350,
            margin=dict(t=40, b=20, l=10, r=10),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)"
        )
        for annotation in fig['layout']['annotations']:
            annotation['font'] = dict(color="#E0E0E0", size=14)

        st.plotly_chart(fig, use_container_width=True)


    draw_unified_donuts(df_mss, "🍩 MSS Composition (Exportable Row)")
    draw_unified_donuts(df_mgw, "🍩 MGW Composition (Exportable Row)")

else:
    st.info("Waiting for data... Please drag and drop a raw alarm file above to generate the dashboard.")