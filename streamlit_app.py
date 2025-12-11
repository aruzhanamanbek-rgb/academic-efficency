import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="KIMEP Academic Load Intelligence", layout="wide")

NEON_CSS = """
<style>
header[data-testid="stHeader"] {
    background-color: transparent !important;
}
[data-testid="stToolbar"] {
    display: none;
}
[data-testid="stDecoration"] {
    display: none;
}
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #3a0ca3 0%, #4361ee 45%, #4cc9f0 100%);
    color: #ffffff;
    font-family: "Segoe UI", Roboto, Arial;
}
[data-testid="stSidebar"] {
    background: linear-gradient(135deg, #3a0ca3 0%, #4361ee 45%, #4cc9f0 100%);
}
[data-testid="stSidebar"] * {
    color: #ffffff !important;
}
[data-testid="stSidebar"] [data-baseweb="select"],
[data-testid="stSidebar"] [data-baseweb="input"],
[data-testid="stSidebar"] .stNumberInput,
[data-testid="stSidebar"] .stSlider {
    background-color: rgba(255, 255, 255, 0.15) !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] [data-baseweb="input"] input,
[data-testid="stSidebar"] .stNumberInput input {
    background-color: rgba(255, 255, 255, 0.2) !important;
    color: #000000 !important;
    font-weight: 500 !important;
}
[data-testid="stSidebar"] .stNumberInput button {
    background-color: rgba(255, 255, 255, 0.3) !important;
    color: #000000 !important;
}
.stDownloadButton button {
    background-color: #4cc9f0 !important;
    color: #000000 !important;
    border: none !important;
    font-weight: 600 !important;
}
.stDownloadButton button:hover {
    background-color: #4895ef !important;
    color: #000000 !important;
}
.header-box {
    text-align:center;
    margin-bottom:8px;
}
.big-title {
    font-size:40px;
    font-weight:800;
    color:#f2e6ff;
    text-shadow:0 2px 12px rgba(0,0,0,0.25);
}
.subtitle {
    text-align:center;
    color:rgba(255,255,255,0.88);
    margin-bottom:12px;
    font-size:14px;
}
.neon-card {
    background: rgba(255,255,255,0.03);
    border-radius:12px;
    padding:10px;
    margin-bottom:10px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.25);
}
.chart-title {
    font-size:15px;
    font-weight:700;
    color:#ffffff;
    margin-bottom:6px;
}
.desc {
    background: rgba(255,255,255,0.02);
    padding:8px;
    border-radius:8px;
    color:#ffffff;
    font-size:13px;
    margin-top:8px;
}
.kpi {
    background: rgba(255,255,255,0.04);
    padding:10px;
    border-radius:10px;
    text-align:center;
    color:#fff;
}
</style>
"""
st.markdown(NEON_CSS, unsafe_allow_html=True)

st.markdown(
    '<div class="header-box"><div class="big-title">🚀 KIMEP Academic Load Intelligence</div>'
    '<div class="subtitle">Academic Efficiency - top 10 focused </div></div>',
    unsafe_allow_html=True
)

@st.cache_data
def load_and_clean(path="cleaned_schedule.xlsx"):
    try:
        df = pd.read_excel(path)
    except Exception:
        return None
    
    df.columns = df.columns.astype(str).str.strip().str.replace(r"\s+", "_", regex=True).str.lower()
    
    expected = {
        "course_title": "", "code": "", "class_dates": "", "class_times": "",
        "days": "", "hall": "", "instructor": "", "minutes": 0,
        "start_time": "", "end_time": "", "department": ""
    }
    for k, v in expected.items():
        if k not in df.columns:
            df[k] = v
    
    df["minutes"] = pd.to_numeric(df["minutes"], errors="coerce").fillna(0).astype(int)
    
    valid_days = {"Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"}
    def normalize_day(raw):
        if pd.isna(raw):
            return np.nan
        s = str(raw).strip()
        s_cap = s.capitalize()
        mapping = {"M": "Mon", "T": "Tue", "W": "Wed", "R": "Thu", "F": "Fri", "S": "Sat", "Su": "Sun"}
        if s_cap in valid_days:
            return s_cap
        if s in mapping:
            return mapping[s]
        return np.nan
    
    df["days_clean"] = df["days"].apply(normalize_day)
    df = df[~df["days_clean"].isna()].copy()
    df["days"] = df["days_clean"]
    df.drop(columns=["days_clean"], inplace=True)
    
    def parse_hour(x):
        try:
            t = pd.to_datetime(str(x), errors="coerce")
            if pd.isna(t):
                return np.nan
            return t.hour + t.minute / 60.0
        except:
            return np.nan
    
    df["start_hour"] = df["start_time"].apply(parse_hour)
    
    df["instructor"] = df["instructor"].astype(str).replace({"nan": ""}).str.strip().replace({"": "Unknown"})
    df["hall"] = df["hall"].astype(str).replace({"nan": ""}).str.strip().replace({"": "Unknown"})
    df["course_title"] = df["course_title"].astype(str).replace({"nan": ""}).str.strip().replace({"": "Unknown"})
    
    def extract_faculty(code_str):
        if pd.isna(code_str) or str(code_str).strip() == "":
            return "Unknown"
        
        code = str(code_str).strip()
        import re
        match = re.match(r'^([A-Z]+)', code)
        if not match:
            return "Unknown"
        
        prefix = match.group(1)
        
        faculty_map = {
            "ACC": "Bang College of Business",
            "BUS": "Bang College of Business",
            "FIN": "Bang College of Business",
            "MGT": "Bang College of Business",
            "MKT": "Bang College of Business",
            "OPM": "Bang College of Business",
            "IFS": "Bang College of Business",
            "EBA": "Bang College of Business",
            "MM": "Bang College of Business",
            "ECN": "College of Social Sciences",
            "IRL": "College of Social Sciences",
            "POL": "College of Social Sciences",
            "PAD": "College of Social Sciences",
            "PAF": "College of Social Sciences",
            "SOC": "College of Social Sciences",
            "CSS": "College of Social Sciences",
            "GEN": "College of Social Sciences",
            "JMC": "College of Human Sciences & Education",
            "PSY": "College of Human Sciences & Education",
            "EPM": "College of Human Sciences & Education",
            "TFL": "College of Human Sciences & Education",
            "TRN": "College of Human Sciences & Education",
            "LING": "College of Human Sciences & Education",
            "COGN": "College of Human Sciences & Education",
            "ENG": "College of Human Sciences & Education",
            "KAZ": "College of Human Sciences & Education",
            "RUS": "College of Human Sciences & Education",
            "CHN": "College of Human Sciences & Education",
            "GER": "College of Human Sciences & Education",
            "KOR": "College of Human Sciences & Education",
            "LDP": "College of Human Sciences & Education",
            "FOP": "College of Human Sciences & Education",
            "LAW": "Law School",
            "CIT": "School of Computer Science & Mathematics",
            "CLP": "School of Computer Science & Mathematics",
            "SCS": "School of Computer Science & Mathematics",
            "MATH": "School of Computer Science & Mathematics",
        }
        
        return faculty_map.get(prefix, "Other")
    
    df["department"] = df["code"].apply(extract_faculty)
    
    return df

df = load_and_clean()
if df is None:
    st.warning("Please upload cleaned_schedule.xlsx")
    uploaded = st.file_uploader("Upload cleaned_schedule.xlsx", type=["xlsx","xls"])
    if not uploaded:
        st.stop()
    df = load_and_clean(uploaded)

st.sidebar.markdown("### 🔎 Filters (leave empty to show all)")

days_options = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
instructors = sorted(df["instructor"].replace("Unknown", np.nan).dropna().unique().tolist())
faculties = sorted(df["department"].replace("Unknown", np.nan).replace("Other", np.nan).dropna().unique().tolist())
halls = sorted(df["hall"].replace("Unknown", np.nan).dropna().unique().tolist())

st.sidebar.markdown("👤 **Instructor**")
inst_sel = st.sidebar.multiselect("Select instructors", options=instructors, default=[], label_visibility="collapsed")

st.sidebar.markdown("🏢 **Faculty**")
dept_sel = st.sidebar.multiselect("Select faculties", options=faculties, default=[], label_visibility="collapsed")

st.sidebar.markdown("🏛️ **Hall**")
hall_sel = st.sidebar.multiselect("Select halls", options=halls, default=[], label_visibility="collapsed")

st.sidebar.markdown("📅 **Days**")
days_sel = st.sidebar.multiselect("Select days", options=days_options, default=[], label_visibility="collapsed")

st.sidebar.markdown("⏰ **Hour range (08:30–21:45)**")
time_range = st.sidebar.slider("Time range", 8.3, 21.45, (8.3, 21.45), step=0.25, label_visibility="collapsed")

top_n = 10

df_f = df.copy()
if inst_sel:
    df_f = df_f[df_f["instructor"].isin(inst_sel)]
if dept_sel:
    df_f = df_f[df_f["department"].isin(dept_sel)]
if hall_sel:
    df_f = df_f[df_f["hall"].isin(hall_sel)]
if days_sel:
    df_f = df_f[df_f["days"].isin(days_sel)]

mask_time = df_f["start_hour"].notna() & (df_f["start_hour"] >= time_range[0]) & (df_f["start_hour"] <= time_range[1])
df_f = pd.concat([df_f[mask_time], df_f[df_f["start_hour"].isna()]]).drop_duplicates()

if df_f.empty:
    st.warning("No rows matched filters — showing full cleaned dataset.")
    df_f = df.copy()

def style_figure(fig):
    fig.update_layout(
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin=dict(l=10, r=10, t=28, b=20),
        font=dict(color="black")
    )
    fig.update_xaxes(showgrid=True, gridcolor="lightgray", showline=True, linecolor="lightgray")
    fig.update_yaxes(showgrid=True, gridcolor="lightgray", showline=True, linecolor="lightgray")
    return fig

k1, k2, k3, k4 = st.columns(4)
k1.markdown(f'<div class="kpi"><b>Sessions</b><br>{len(df_f)}</div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi"><b>Unique Courses</b><br>{df_f["course_title"].nunique()}</div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi"><b>Active Instructors</b><br>{df_f["instructor"].nunique()}</div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi"><b>Total Hours</b><br>{df_f["minutes"].sum()/60:.1f} h</div>', unsafe_allow_html=True)

st.markdown('<div class="neon-card">', unsafe_allow_html=True)

avg_minutes_per_instructor = df_f.groupby("instructor")["minutes"].sum().mean()
top_instructor = df_f.groupby("instructor")["minutes"].sum().sort_values(ascending=False).head(1)
top_instructor_name = top_instructor.index[0] if len(top_instructor) > 0 else "N/A"
top_instructor_hours = top_instructor.values[0] / 60 if len(top_instructor) > 0 else 0
percentage_above_avg = ((top_instructor_hours - avg_minutes_per_instructor/60) / (avg_minutes_per_instructor/60) * 100) if avg_minutes_per_instructor > 0 else 0

if not df_f[df_f["start_hour"].notna()].empty:
    df_peak = df_f[df_f["start_hour"].notna()].copy()
    df_peak["hour_slot"] = (df_peak["start_hour"] // 1).astype(int)
    peak_analysis = df_peak.groupby(["days", "hour_slot"]).size().sort_values(ascending=False).head(1)
    if len(peak_analysis) > 0:
        peak_day = peak_analysis.index[0][0]
        peak_hour = int(peak_analysis.index[0][1])
        peak_count = peak_analysis.values[0]
        peak_text = f"{peak_day} {peak_hour:02d}:00-{peak_hour+1:02d}:00 ({peak_count} sessions)"
    else:
        peak_text = "N/A"
else:
    peak_text = "N/A"

top_hall = df_f.groupby("hall")["minutes"].sum().sort_values(ascending=False).head(1)
top_hall_name = top_hall.index[0] if len(top_hall) > 0 else "N/A"
top_hall_hours = top_hall.values[0] / 60 if len(top_hall) > 0 else 0

c_a, c_b, c_c = st.columns(3)

with c_a:
    st.markdown(f"""
    <div style='background: rgba(76, 201, 240, 0.15); padding: 12px; border-radius: 10px; border-left: 4px solid #4cc9f0;'>
    <div style='font-size: 13px; color: #4cc9f0; font-weight: 600; margin-bottom: 4px;'>📊 DATA COVERAGE</div>
    <div style='font-size: 12px; color: #fff; line-height: 1.6;'>
    • <b>{len(df)}</b> total sessions analyzed<br>
    • <b>{df['department'].nunique()}</b> faculties covered<br>
    • <b>{df['course_title'].nunique()}</b> unique courses<br>
    • <b>{df['instructor'].nunique()}</b> instructors tracked
    </div>
    </div>
    """, unsafe_allow_html=True)

with c_b:
    st.markdown(f"""
    <div style='background: rgba(181, 23, 158, 0.15); padding: 12px; border-radius: 10px; border-left: 4px solid #b5179e;'>
    <div style='font-size: 13px; color: #f72585; font-weight: 600; margin-bottom: 4px;'>⚡ WORKLOAD INSIGHT</div>
    <div style='font-size: 12px; color: #fff; line-height: 1.6;'>
    • Top instructor: <b>{top_instructor_name}</b><br>
    • Teaching load: <b>{top_instructor_hours:.1f} hours/week</b><br>
    • <b>{percentage_above_avg:+.0f}%</b> vs. average<br>
    • Suggests workload rebalancing
    </div>
    </div>
    """, unsafe_allow_html=True)

with c_c:
    st.markdown(f"""
    <div style='background: rgba(242, 117, 133, 0.15); padding: 12px; border-radius: 10px; border-left: 4px solid #f27585;'>
    <div style='font-size: 13px; color: #f27585; font-weight: 600; margin-bottom: 4px;'>🕐 UTILIZATION PEAK</div>
    <div style='font-size: 12px; color: #fff; line-height: 1.6;'>
    • Peak time slot: <b>{peak_text}</b><br>
    • Most used hall: <b>{top_hall_name}</b><br>
    • Hall usage: <b>{top_hall_hours:.1f} hours/week</b><br>
    • Consider load spreading
    </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

CH_H = 320
NEON = ["#4cc9f0", "#7209b7", "#4895ef", "#560bad", "#b5179e", "#f72585", "#3f37c9"]

c1, c2 = st.columns(2)

with c1:
    st.markdown('<div class="neon-card"><div class="chart-title">🟣 Top Instructors - Load (Top 10)</div>', unsafe_allow_html=True)
    inst_load = df_f.groupby("instructor", as_index=False)["minutes"].sum().sort_values("minutes", ascending=False).head(10)
    if not inst_load.empty:
        fig = px.scatter(inst_load, x="instructor", y="minutes", size="minutes", color="minutes",
                        color_continuous_scale=[NEON[0], NEON[1]],
                        labels={"minutes":"Minutes","instructor":"Instructor"}, height=CH_H)
        fig.update_traces(marker=dict(line=dict(width=0.4, color="rgba(0,0,0,0.1)")))
        fig = style_figure(fig)
        fig.update_xaxes(tickangle=-35, tickfont=dict(size=10))
        fig.update_yaxes(title_text="Total minutes")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No instructor data.")
    st.markdown('<div class="desc">• Bubble size = total teaching minutes per instructor. • Top 10 prevents overlap. Hover for exact values.</div></div>', unsafe_allow_html=True)

with c2:
    st.markdown('<div class="neon-card"><div class="chart-title">🏛️ Hall Usage - Top Rooms (Top 10)</div>', unsafe_allow_html=True)
    hall_usage = df_f.groupby("hall", as_index=False)["minutes"].sum().sort_values("minutes", ascending=False).head(10)
    if not hall_usage.empty:
        fig = px.bar(hall_usage, x="minutes", y="hall", orientation="h", color="minutes",
                    color_continuous_scale=[NEON[2], NEON[0]],
                    labels={"minutes":"Minutes","hall":"Hall"}, height=CH_H)
        fig.update_traces(marker_line_color="rgba(0,0,0,0.1)")
        fig = style_figure(fig)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hall data.")
    st.markdown('<div class="desc">• Horizontal bars show most used halls. Useful for allocation planning.</div></div>', unsafe_allow_html=True)

c3, c4 = st.columns(2)

with c3:
    st.markdown('<div class="neon-card"><div class="chart-title">🔥 Weekly Intensity - Heatmap (hour bins)</div>', unsafe_allow_html=True)
    df_h = df_f[df_f["start_hour"].notna()].copy()
    if not df_h.empty:
        bins = np.arange(8.5, 22.0, 1.0)
        df_h["hour_bin"] = pd.cut(df_h["start_hour"], bins=bins, include_lowest=True).astype(str)
        heat = df_h.groupby(["days", "hour_bin"]).size().reset_index(name="count")
        if not heat.empty:
            fig = px.density_heatmap(heat, x="hour_bin", y="days", z="count",
                                    labels={"hour_bin":"Start hour bin","days":"Day","count":"Sessions"},
                                    color_continuous_scale="Viridis", height=CH_H)
            fig = style_figure(fig)
            fig.update_xaxes(tickangle=-45, tickfont=dict(size=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Insufficient heatmap data.")
    else:
        st.info("No parsed start_time available for heatmap.")
    st.markdown('<div class="desc">• Heatmap highlights busiest day/hour windows. Binned by hour for clarity.</div></div>', unsafe_allow_html=True)

with c4:
    st.markdown('<div class="neon-card"><div class="chart-title">🌞 Faculty → Instructor - Sunburst (Top 5)</div>', unsafe_allow_html=True)
    dept_tot = df_f.groupby("department", as_index=False)["minutes"].sum().sort_values("minutes", ascending=False).head(5)
    top_depts = dept_tot["department"].tolist()
    
    sb_df = df_f[df_f["department"].isin(top_depts)].groupby(["department","instructor"], as_index=False)["minutes"].sum()
    
    def get_last_name(full_name):
        if pd.isna(full_name) or str(full_name).strip() == "":
            return "Unknown"
        name = str(full_name).strip()
        if "," in name:
            return name.split(",")[0].strip()
        else:
            parts = name.split()
            return parts[-1] if parts else name
    
    sb_df["instructor_short"] = sb_df["instructor"].apply(get_last_name)
    
    sb_small = []
    for d in top_depts:
        tmp = sb_df[sb_df["department"]==d].sort_values("minutes", ascending=False).head(3)
        sb_small.append(tmp)
    
    if sb_small:
        sb_final = pd.concat(sb_small, ignore_index=True)
        fig = px.sunburst(sb_final, path=["department","instructor_short"], values="minutes",
                         height=CH_H, color="minutes", color_continuous_scale="Ice")
        fig = style_figure(fig)
        fig.update_traces(textfont=dict(size=12))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Not enough data for sunburst.")
    st.markdown('<div class="desc">• Sunburst shows faculty workload across top instructors (top 5 faculties, top 3 instructors per faculty).</div></div>', unsafe_allow_html=True)

c5, c6 = st.columns(2)

with c5:
    st.markdown('<div class="neon-card"><div class="chart-title">📚 Top Frequent Courses (by sessions)</div>', unsafe_allow_html=True)
    popular = df_f["course_title"].value_counts().head(10).reset_index()
    popular.columns = ["course_title","count"]
    if not popular.empty:
        fig = px.bar(popular, x="count", y="course_title", orientation="h", height=CH_H,
                    color_discrete_sequence=[NEON[4]])
        fig = style_figure(fig)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No course frequency data.")
    st.markdown('<div class="desc">• Shows most frequently scheduled courses (top 10) to avoid label crowding.</div></div>', unsafe_allow_html=True)

with c6:
    st.markdown('<div class="neon-card"><div class="chart-title">⏳ Courses by Total Minutes</div>', unsafe_allow_html=True)
    course_min = df_f.groupby("course_title", as_index=False)["minutes"].sum().sort_values("minutes", ascending=False).head(10)
    if not course_min.empty:
        fig = px.bar(course_min, x="minutes", y="course_title", orientation="h", height=CH_H,
                    color="minutes", color_continuous_scale=[NEON[1], NEON[0]])
        fig = style_figure(fig)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No minutes-per-course data.")
    st.markdown('<div class="desc">• Aggregated minutes per course (top 10) helps find heavy courses.</div></div>', unsafe_allow_html=True)

st.markdown('<div class="neon-card"><div class="chart-title">📊 Faculty Distribution - Minutes</div>', unsafe_allow_html=True)
dept = df_f[~df_f["department"].isin(["Other", "Unknown"])].groupby("department", as_index=False)["minutes"].sum().sort_values("minutes", ascending=False).head(10)
if not dept.empty:
    fig = px.pie(dept, names="department", values="minutes", hole=0.35, height=360,
                color_discrete_sequence=NEON)
    fig = style_figure(fig)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No faculty data.")
st.markdown('<div class="desc">• Faculty share limited to top 10 to keep chart readable.</div></div>', unsafe_allow_html=True)

with st.expander("Tips & UX — How to read this dashboard"):
    st.markdown("""
- Leave filters empty to view global statistics.
- Top 10 limit is applied to reduce overlap and keep visualizations legible.
- Heatmap uses hour bins (1h) to show student flow peaks.
- Sunburst is intentionally limited so labels remain readable.
""")

st.markdown("---")
st.markdown('<div class="neon-card">', unsafe_allow_html=True)
st.markdown('<div class="chart-title">📊 Academic Efficiency - Key Insights</div>', unsafe_allow_html=True)

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("""
    <div class="desc">
    <b> Resource Utilization</b><br>
    This dashboard analyzes KIMEP's academic load distribution to identify efficiency patterns:
    <ul>
    <li><b>Instructor workload balance:</b> Bubble chart reveals which instructors carry the heaviest teaching loads, enabling better workload distribution</li>
    <li><b>Classroom optimization:</b> Hall usage bars show which rooms are most utilized, helping optimize space allocation</li>
    <li><b>Peak hour identification:</b> Heatmap highlights when classes cluster, revealing potential scheduling conflicts or underutilized time slots</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

with col_b:
    st.markdown("""
    <div class="desc">
    <b> Strategic Recommendations</b><br>
    Based on the visualizations, academic administrators can:
    <ul>
    <li><b>Balance teaching loads:</b> Redistribute courses among instructors to prevent burnout and ensure quality</li>
    <li><b>Optimize scheduling:</b> Use heatmap insights to spread classes more evenly across the week</li>
    <li><b>Maximize space usage:</b> Allocate high-demand rooms efficiently based on actual utilization data</li>
    <li><b>Faculty planning:</b> Pie chart shows faculty resource allocation for budget planning</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<div style='color:#ffffff;font-size:13px'>Hover on any element to see exact counts/minutes. Download filtered data below.</div>", unsafe_allow_html=True)

csv = df_f.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download filtered CSV", csv, "filtered_schedule_filtered.csv", "text/csv")
