import base64
import os
import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================
# 1. Page Configuration & Custom CSS
# ==========================================
st.set_page_config(
    page_title="ES ISO 9001:2015/2026 Quality Management Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main .block-container { 
        padding-top: 1.2rem; 
        padding-bottom: 2rem; 
        max-width: 95%; 
    }
    .metric-card {
        background-color: #f8f9fa; 
        border-radius: 8px; 
        padding: 12px 16px;
        border-left: 4px solid #1e3d59; 
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .metric-label { 
        font-size: 0.82rem; 
        color: #555555; 
        font-weight: 600; 
        text-transform: uppercase; 
    }
    .metric-value { 
        font-size: 1.55rem; 
        font-weight: 700; 
        color: #1e3d59; 
    }
    .section-card {
        background-color: #ffffff; 
        border: 1px solid #e0e0e0; 
        border-radius: 8px;
        padding: 18px; 
        margin-bottom: 16px; 
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .section-title {
        font-size: 1.05rem; 
        font-weight: 700; 
        color: #1e3d59; 
        margin-bottom: 12px;
        border-bottom: 2px solid #e2e8f0; 
        padding-bottom: 6px;
    }
    /* Warning Box Custom CSS */
    .warning-banner {
        background-color: #fef2f2;
        border: 2px solid #ef4444;
        border-left: 6px solid #dc2626;
        border-radius: 8px;
        padding: 14px 18px;
        margin-top: 16px;
        margin-bottom: 12px;
        box-shadow: 0 2px 4px rgba(220, 38, 38, 0.08);
    }
    .warning-title {
        color: #991b1b;
        font-size: 1.05rem;
        font-weight: 800;
        letter-spacing: 0.5px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .warning-text {
        color: #7f1d1d;
        font-size: 0.90rem;
        margin-top: 4px;
        font-weight: 500;
    }
    /* Logo Image Rule */
    .header-logo-img {
        height: 120px !important;
        width: auto !important;
        object-fit: contain;
        display: block;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

EXCEL_FILE = "QMS_Data.xlsx"


# ==========================================
# 2. Multi-Sheet Data Loader
# ==========================================
@st.cache_data(ttl=5)
def load_excel_data():
  if not os.path.exists(EXCEL_FILE):
    return None
  xls = pd.ExcelFile(EXCEL_FILE)
  df_kpi = pd.read_excel(xls, "Overview_KPIs")
  df_rc = pd.read_excel(xls, "Deviations_RootCauses")

  df_defn = (
      pd.read_excel(xls, "Defn of ACs")
      if "Defn of ACs" in xls.sheet_names
      else pd.DataFrame()
  )

  df_audits = pd.read_excel(xls, "Audits_Compliance")
  df_actions = pd.read_excel(xls, "Operational_Actions")
  df_sec = pd.read_excel(xls, "Secondary_Metrics")

  df_merged = df_kpi.merge(
      df_rc, on=["Site", "Department", "Period"], how="left"
  )
  df_merged = df_merged.merge(
      df_audits, on=["Site", "Department", "Period"], how="left"
  )
  df_merged = df_merged.merge(
      df_actions, on=["Site", "Department", "Period"], how="left"
  )
  df_merged = df_merged.merge(
      df_sec, on=["Site", "Department", "Period"], how="left"
  )

  return {
      "merged": df_merged,
      "kpi": df_kpi,
      "rc": df_rc,
      "defn": df_defn,
      "audits": df_audits,
      "actions": df_actions,
      "sec": df_sec,
  }


data_dict = load_excel_data()

if data_dict is None:
  st.error(
      f"File '{EXCEL_FILE}' not found. Please run 'py generate_excel.py' in your"
      " terminal first."
  )
  st.stop()

df = data_dict["merged"]

# ==========================================
# 3. Sidebar Filters & Navigation
# ==========================================
st.sidebar.title("🔍 Navigation & Filters")

sites = df["Site"].dropna().unique().tolist()
selected_site = st.sidebar.selectbox("Select Site", sites, index=0)

depts = (
    df[df["Site"] == selected_site]["Department"].dropna().unique().tolist()
)
selected_dept = st.sidebar.selectbox("Select Department", depts, index=0)

periods = (
    df[(df["Site"] == selected_site) & (df["Department"] == selected_dept)][
        "Period"
    ]
    .dropna()
    .unique()
    .tolist()
)
selected_period = st.sidebar.selectbox(
    "Select Reporting Period", periods, index=0
)

filtered = df[
    (df["Site"] == selected_site)
    & (df["Department"] == selected_dept)
    & (df["Period"] == selected_period)
]

if filtered.empty:
  st.warning("No record found for selected key combination.")
  st.stop()

row = filtered.iloc[0]

# ==========================================
# 4. Top Header with Precision Vertical Offset Alignment
# ==========================================
col_logo, col_title = st.columns([1.8, 5.2])

with col_logo:
  logo_file = None
  for ext in ["png", "jpg", "jpeg"]:
    if os.path.exists(f"logo.{ext}"):
      logo_file = f"logo.{ext}"
      break

  if logo_file:
    with open(logo_file, "rb") as f:
      encoded_logo = base64.b64encode(f.read()).decode()
    st.markdown(
        f'<img src="data:image/png;base64,{encoded_logo}"'
        ' class="header-logo-img">',
        unsafe_allow_html=True,
    )
  else:
    st.markdown(
        """
        <div style="
            height: 120px; 
            background-color: #1e3d59; 
            color: #ffffff; 
            border-radius: 6px; 
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            font-weight: bold;
            font-size: 1.1rem;">
            INSTITUTE LOGO
            <span style="font-size: 0.75rem; font-weight: normal; color: #cbd5e1;">(logo.png missing)</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_title:
  st.markdown(
      f"""
        <div style="
            display: flex; 
            flex-direction: column; 
            justify-content: space-between; 
            height: 120px; 
            box-sizing: border-box;
            padding: 0px;
            margin: 0px;">
            <!-- Top Line: Standard Badge (Offset slightly below top edge) -->
            <div style="line-height: 1; margin-top: 6px; padding: 0;">
                <span style="
                    background-color: #1e3d59; 
                    color: #ffffff; 
                    padding: 4px 10px; 
                    border-radius: 4px; 
                    font-size: 0.90rem; 
                    font-weight: 700; 
                    letter-spacing: 0.5px; 
                    display: inline-block;
                    line-height: 1;">
                    ES ISO 9001:2015/2026
                </span>
            </div>
            <!-- Middle Line: Title Header -->
            <div style="
                color: #1e3d59; 
                font-size: 2.2rem; 
                font-weight: 800; 
                line-height: 1.05; 
                letter-spacing: -0.5px;
                margin: 0;
                padding: 0;">
                Quality Management Dashboard
            </div>
            <!-- Bottom Line: Metadata (Strictly aligned to lower edge) -->
            <div style="
                color: #64748b; 
                font-size: 0.90rem; 
                font-weight: 500; 
                line-height: 1;
                margin-bottom: 2px;
                padding: 0;">
                Data Source: <span style="color: #334155; font-weight: 600;">{EXCEL_FILE}</span> &nbsp;|&nbsp; 
                Site: <span style="color: #334155; font-weight: 600;">{selected_site}</span> &nbsp;|&nbsp; 
                Department: <span style="color: #334155; font-weight: 600;">{selected_dept}</span> &nbsp;|&nbsp; 
                Period: <span style="color: #334155; font-weight: 600;">{selected_period}</span>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )

st.markdown("---")

# ==========================================
# 5. Core Overview KPI Cards
# ==========================================
c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
  st.markdown(
      f'<div class="metric-card"><div class="metric-label">Total'
      f' Deviations</div><div'
      f' class="metric-value">{row.get("Total_Deviations", 0)}</div></div>',
      unsafe_allow_html=True,
  )
with c2:
  st.markdown(
      f'<div class="metric-card" style="border-left-color: #2e7d32;"><div'
      ' class="metric-label">Closed On Time</div><div'
      f' class="metric-value">{row.get("Closed_On_Time", 0)}</div></div>',
      unsafe_allow_html=True,
  )
with c3:
  st.markdown(
      f'<div class="metric-card" style="border-left-color: #c62828;"><div'
      ' class="metric-label">Overdue</div><div'
      f' class="metric-value">{row.get("Overdue", 0)}</div></div>',
      unsafe_allow_html=True,
  )
with c4:
  st.markdown(
      f'<div class="metric-card" style="border-left-color: #1565c0;"><div'
      ' class="metric-label">CAPA Closed</div><div'
      f' class="metric-value">{row.get("CAPA_Closed", 0)}</div></div>',
      unsafe_allow_html=True,
  )
with c5:
  st.markdown(
      f'<div class="metric-card" style="border-left-color: #d84315;"><div'
      ' class="metric-label">CAPA Failed</div><div'
      f' class="metric-value">{row.get("CAPA_Failed", 0)}</div></div>',
      unsafe_allow_html=True,
  )
with c6:
  st.markdown(
      f'<div class="metric-card" style="border-left-color: #6a1b9a;"><div'
      ' class="metric-label">Training Comp.</div><div'
      f' class="metric-value">{row.get("Training_Compliance", 0)}%</div></div>',
      unsafe_allow_html=True,
  )

# ==========================================
# 5.1 Dynamic Warning Alert Banner
# ==========================================
overdue_val = int(row.get("Overdue", 0))
capa_failed_val = int(row.get("CAPA_Failed", 0))

if overdue_val > 0 or capa_failed_val > 0:
  st.markdown(
      f"""
        <div class="warning-banner">
            <div class="warning-title">
                ⚠️ CRITICAL COMPLIANCE WARNING: OVERDUE DEVIATIONS & FAILED CAPAs DETECTED
            </div>
            <div class="warning-text">
                Attention required for <b>{selected_dept}</b> ({selected_site}) during period <b>{selected_period}</b>:
                <ul style="margin-top: 6px; margin-bottom: 0px; padding-left: 20px;">
                    {"<li><b>Overdue Deviations:</b> <span style='color: #dc2626; font-weight: 700;'>" + str(overdue_val) + " item(s)</span> past target resolution time limits.</li>" if overdue_val > 0 else ""}
                    {"<li><b>Failed CAPAs:</b> <span style='color: #dc2626; font-weight: 700;'>" + str(capa_failed_val) + " item(s)</span> requiring immediate root cause re-evaluation and escalation.</li>" if capa_failed_val > 0 else ""}
                </ul>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )
else:
  st.markdown(
      """
        <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-left: 5px solid #16a34a; border-radius: 6px; padding: 10px 16px; margin-top: 14px; margin-bottom: 10px;">
            <span style="color: #15803d; font-weight: 700; font-size: 0.90rem;">
                ✅ COMPLIANCE STATUS OPTIMAL: No overdue deviations or failed CAPAs reported for this period.
            </span>
        </div>
        """,
      unsafe_allow_html=True,
  )

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 6. Trend and Root Cause Charts
# ==========================================
left_col, right_col = st.columns([1, 1])

with left_col:
  st.markdown('<div class="section-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="section-title">📈 Historical Deviation Trend</div>',
      unsafe_allow_html=True,
  )
  trend_df = df[
      (df["Site"] == selected_site) & (df["Department"] == selected_dept)
  ]
  fig_trend = px.line(
      trend_df,
      x="Period",
      y="Total_Deviations",
      markers=True,
      labels={"Total_Deviations": "Deviations", "Period": "Period"},
  )
  fig_trend.update_traces(line_color="#1e3d59", line_width=3, marker_size=8)
  fig_trend.update_layout(
      height=300, margin=dict(l=20, r=20, t=30, b=20), hovermode="x"
  )
  st.plotly_chart(fig_trend, use_container_width=True)
  st.markdown("</div>", unsafe_allow_html=True)

with right_col:
  st.markdown('<div class="section-card">', unsafe_allow_html=True)
  st.markdown(
      '<div class="section-title">🧩 Root Cause Breakdown</div>',
      unsafe_allow_html=True,
  )

  rc_df = pd.DataFrame({
      "Category": ["People", "Process", "System", "Equipment", "Other"],
      "Count": [
          row.get("RC_People", 0),
          row.get("RC_Process", 0),
          row.get("RC_System", 0),
          row.get("RC_Equipment", 0),
          row.get("RC_Other", 0),
      ],
  })

  fig_rc = px.pie(
      rc_df,
      names="Category",
      values="Count",
      hole=0.4,
      color="Category",
      color_discrete_map={
          "People": "#00B050",
          "Process": "#ED7D31",
          "System": "#5B9BD5",
          "Equipment": "#FF66C4",
          "Other": "#70AD47",
      },
  )
  fig_rc.update_layout(height=260, margin=dict(l=20, r=20, t=10, b=10))
  st.plotly_chart(fig_rc, use_container_width=True)

  with st.expander("ℹ️ View Root Cause Category Definitions"):
    rc_definitions = [
        {
            "title": "1. RC_People (Human Factors & Personnel)",
            "bg": "#00B050",
            "text_color": "#ffffff",
            "def": (
                "Deviations originating from human error, procedural"
                " non-compliance, training deficits, oversight, fatigue, or gaps"
                " in technical competency."
            ),
            "context": (
                "Identifies areas requiring refresher training, standard"
                " operating procedure (SOP) reassessment, or improved"
                " operational supervision."
            ),
        },
        {
            "title": (
                "2. RC_Process (Operational & Methodological Procedures)"
            ),
            "bg": "#ED7D31",
            "text_color": "#ffffff",
            "def": (
                "Issues stemming from flawed, ambiguous, outdated, or overly"
                " complex workflows, execution protocols, and standard"
                " operational procedures."
            ),
            "context": (
                "Highlights standard operating procedures that need revision,"
                " simplification, or structural alignment to meet quality"
                " standards."
            ),
        },
        {
            "title": (
                "3. RC_System (Management Systems & Information Technology)"
            ),
            "bg": "#5B9BD5",
            "text_color": "#ffffff",
            "def": (
                "Failures within broader governance structures, document"
                " control networks, digital traceability architectures (e.g., QR"
                " code serialization/GTIN databases), software glitches, or"
                " cross-departmental communication gaps."
            ),
            "context": (
                "Focuses on upgrading digital tools, IT workflows, quality"
                " records systems, and management control systems."
            ),
        },
        {
            "title": (
                "4. RC_Equipment (Hardware, Instruments & Infrastructure)"
            ),
            "bg": "#FF66C4",
            "text_color": "#ffffff",
            "def": (
                "Deviations caused by physical machinery, testing laboratory"
                " equipment calibration drift, tool degradation, unexpected"
                " hardware breakdowns, or utility interruptions."
            ),
            "context": (
                "Points to required preventative maintenance schedules,"
                " recalibrations, equipment upgrades, or repair protocols."
            ),
        },
        {
            "title": "5. RC_Other (External & Unclassified Factors)",
            "bg": "#70AD47",
            "text_color": "#ffffff",
            "def": (
                "Root causes that do not fall under traditional operational"
                " categories, including raw material/supplier anomalies,"
                " environmental variations (e.g., temperature/humidity"
                " spikes), or external third-party disruptions."
            ),
            "context": (
                "Triggers supplier quality audits, environmental control"
                " reviews, or external risk mitigation strategies."
            ),
        },
    ]

    for item in rc_definitions:
      st.markdown(
          f"""
            <div style="margin-bottom: 12px; border: 1px solid #e0e0e0; border-radius: 6px; overflow: hidden;">
                <div style="background-color: {item['bg']}; color: {item['text_color']}; padding: 6px 12px; font-weight: bold; font-size: 0.95rem;">
                    {item['title']}
                </div>
                <div style="padding: 8px 12px; background-color: #ffffff; font-size: 0.85rem; color: #333333;">
                    <p style="margin: 0 0 4px 0;"><b>Definition:</b> {item['def']}</p>
                    <p style="margin: 0;"><b>QMS Context:</b> {item['context']}</p>
                </div>
            </div>
            """,
          unsafe_allow_html=True,
      )

  st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 7. Audits & Secondary Metrics
# ==========================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">📋 Audits & Performance Indicators</div>',
    unsafe_allow_html=True,
)

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
  st.metric(
      "Investigation Effectiveness",
      f"{row.get('Investigation_Effectiveness', 0)}%",
  )
with m2:
  st.metric("Total CAPAs", row.get("Total_CAPAs", 0))
with m3:
  st.metric("Open Changes", row.get("Open_Changes", 0))
with m4:
  st.metric("Doc Up-To-Date", f"{row.get('Doc_Up_To_Date', 0)}%")
with m5:
  st.metric("Complaints Received", row.get("Complaints_Received", 0))

st.markdown("---")
a1, a2, a3 = st.columns(3)
with a1:
  st.caption("Internal Audits")
  st.write(
      f"Conducted: **{row.get('Audit_Internal_Conducted', 0)}** | Observations:"
      f" **{row.get('Audit_Internal_Obs', 0)}**"
  )
with a2:
  st.caption("Supplier Audits")
  st.write(
      f"Conducted: **{row.get('Audit_Supplier_Conducted', 0)}** | Observations:"
      f" **{row.get('Audit_Supplier_Obs', 0)}**"
  )
with a3:
  st.caption("Regulatory Audits")
  st.write(
      f"Conducted: **{row.get('Audit_Regulatory_Conducted', 0)}** |"
      f" Observations: **{row.get('Audit_Regulatory_Obs', 0)}**"
  )
st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 8. Operational Management Form
# ==========================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">📝 Operational Management: Actions & Future'
    " Plans</div>",
    unsafe_allow_html=True,
)

with st.form("edit_actions_form"):
  c_act1, c_act2, c_act3 = st.columns(3)
  with c_act1:
    act_in = st.text_area(
        "📌 Actions Taken", value=str(row.get("Actions_Taken", "")), height=140
    )
  with c_act2:
    plan_in = st.text_area(
        "🚀 Future Plans", value=str(row.get("Future_Plans", "")), height=140
    )
  with c_act3:
    note_in = st.text_area(
        "💡 Executive Notes",
        value=str(row.get("Executive_Notes", "")),
        height=140,
    )

  save_btn = st.form_submit_button("💾 Save Updates to Excel Sheet")

  if save_btn:
    actions_df = data_dict["actions"]
    idx = actions_df[
        (actions_df["Site"] == selected_site)
        & (actions_df["Department"] == selected_dept)
        & (actions_df["Period"] == selected_period)
    ].index

    if not idx.empty:
      actions_df.loc[idx, "Actions_Taken"] = act_in
      actions_df.loc[idx, "Future_Plans"] = plan_in
      actions_df.loc[idx, "Executive_Notes"] = note_in

      with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        data_dict["kpi"].to_excel(
            writer, sheet_name="Overview_KPIs", index=False
        )
        data_dict["rc"].to_excel(
            writer, sheet_name="Deviations_RootCauses", index=False
        )
        if not data_dict["defn"].empty:
          data_dict["defn"].to_excel(
              writer, sheet_name="Defn of ACs", index=False
          )
        data_dict["audits"].to_excel(
            writer, sheet_name="Audits_Compliance", index=False
        )
        actions_df.to_excel(
            writer, sheet_name="Operational_Actions", index=False
        )
        data_dict["sec"].to_excel(
            writer, sheet_name="Secondary_Metrics", index=False
        )

      st.success("Updated 'Operational_Actions' sheet in Excel successfully!")
      st.cache_data.clear()
      st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 9. Excel Multi-Sheet Inspector
# ==========================================
st.markdown("---")
with st.expander("📂 View Classified Excel Sheets (Data Inspection)"):
  tab1, tab2, tab_defn, tab3, tab4, tab5 = st.tabs([
      "Overview_KPIs",
      "Deviations_RootCauses",
      "Defn of ACs",
      "Audits_Compliance",
      "Operational_Actions",
      "Secondary_Metrics",
  ])
  with tab1:
    st.dataframe(data_dict["kpi"], use_container_width=True)
  with tab2:
    st.dataframe(data_dict["rc"], use_container_width=True)
  with tab_defn:

    def color_rows(val):
      s = str(val)
      if "1. RC_People" in s:
        return "background-color: #00B050; color: white; font-weight: bold;"
      elif "2. RC_Process" in s:
        return "background-color: #ED7D31; color: white; font-weight: bold;"
      elif "3. RC_System" in s:
        return "background-color: #5B9BD5; color: white; font-weight: bold;"
      elif "4. RC_Equipment" in s:
        return "background-color: #FF66C4; color: white; font-weight: bold;"
      elif "5. RC_Other" in s:
        return "background-color: #70AD47; color: white; font-weight: bold;"
      return ""

    if not data_dict["defn"].empty:
      styled_df = data_dict["defn"].style.map(color_rows, subset=["Category"])
      st.dataframe(styled_df, use_container_width=True)
    else:
      st.info("No definition records found.")

  with tab3:
    st.dataframe(data_dict["audits"], use_container_width=True)
  with tab4:
    st.dataframe(data_dict["actions"], use_container_width=True)
  with tab5:
    st.dataframe(data_dict["sec"], use_container_width=True)