import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Load the Excel file
file_path = "Sample_Primavera_Data.xlsx"
df = pd.read_excel(file_path, sheet_name="Project Data")

# Convert date columns
df["Baseline Start"] = pd.to_datetime(df["Baseline Start"])
df["Baseline Finish"] = pd.to_datetime(df["Baseline Finish"])
df["Actual Start"] = pd.to_datetime(df["Actual Start"])
df["Actual Finish"] = pd.to_datetime(df["Actual Finish"])

# ---------------- Gantt Chart ----------------
gantt_fig = px.timeline(
    df,
    x_start="Actual Start",
    x_end="Actual Finish",
    y="Activity Name",
    color="WBS",
    hover_data=["Activity ID", "Activity Code", "Planned %", "Actual %", "Remarks"]
)
gantt_fig.update_layout(
    title="Gantt Chart - Actual Schedule",
    xaxis_title="Date",
    yaxis_title="Activities",
    yaxis_autorange="reversed",
    height=600
)

# ---------------- Milestone Trend Analysis ----------------
milestones = df.iloc[::5].copy()
milestones["Month"] = milestones["Actual Finish"].dt.to_period("M").astype(str)
milestone_fig = px.line(
    milestones,
    x="Month",
    y="Actual Finish",
    color="Activity Name",
    markers=True,
    title="Milestone Trend Analysis"
)

# ---------------- S-Curve ----------------
df_sorted = df.sort_values("Actual Finish")
df_sorted["Cumulative Planned Cost"] = df_sorted["Budgeted Cost"].cumsum()
df_sorted["Cumulative Actual Cost"] = df_sorted["Actual Cost"].cumsum()
s_curve_fig = go.Figure()
s_curve_fig.add_trace(go.Scatter(
    x=df_sorted["Actual Finish"], y=df_sorted["Cumulative Planned Cost"],
    mode='lines+markers', name='Planned Cost'))
s_curve_fig.add_trace(go.Scatter(
    x=df_sorted["Actual Finish"], y=df_sorted["Cumulative Actual Cost"],
    mode='lines+markers', name='Actual Cost'))
s_curve_fig.update_layout(title="S-Curve - Planned vs Actual Cost", xaxis_title="Date", yaxis_title="Cumulative Cost")

# ---------------- Histogram of Man-hours ----------------
hist_fig = px.bar(
    df,
    x="Activity Name",
    y=["Budgeted Hours", "Actual Hours"],
    barmode="group",
    title="Man-hours Histogram",
    labels={"value": "Hours", "variable": "Type"}
)

# ---------------- Earned Value Management ----------------
df["EV"] = df["Budgeted Cost"] * df["Actual %"] / 100
df["PV"] = df["Budgeted Cost"] * df["Planned %"] / 100
df["AC"] = df["Actual Cost"]
EV = df["EV"].sum()
PV = df["PV"].sum()
AC = df["AC"].sum()
CPI = EV / AC if AC else None
SPI = EV / PV if PV else None
CV = EV - AC
SV = EV - PV
AT = 100
ES = SPI * AT if SPI else None
Schedule_Variance_Time = ES - AT if ES else None

summary_fig = go.Figure(data=[
    go.Table(header=dict(values=["Metric", "Value"]),
             cells=dict(values=[
                 ["Planned Value (PV)", "Earned Value (EV)", "Actual Cost (AC)",
                  "Cost Variance (CV)", "Schedule Variance (SV)",
                  "Cost Performance Index (CPI)", "Schedule Performance Index (SPI)",
                  "Earned Schedule (ES)", "Schedule Variance (Time)"],
                 [f"${PV:,.2f}", f"${EV:,.2f}", f"${AC:,.2f}",
                  f"${CV:,.2f}", f"${SV:,.2f}", 
                  f"{CPI:.2f}" if CPI else "N/A", f"{SPI:.2f}" if SPI else "N/A",
                  f"{ES:.2f} days" if ES else "N/A", f"{Schedule_Variance_Time:.2f} days" if Schedule_Variance_Time else "N/A"]
             ]))
])
summary_fig.update_layout(title="Earned Value and Earned Schedule Summary")

# ---------------- CPI/SPI Trend Chart ----------------
df_grouped = df.groupby("Actual Finish").agg({
    "EV": "sum", "PV": "sum", "AC": "sum"
}).sort_index().reset_index()
df_grouped["CPI"] = df_grouped["EV"] / df_grouped["AC"]
df_grouped["SPI"] = df_grouped["EV"] / df_grouped["PV"]

trend_fig = go.Figure()
trend_fig.add_trace(go.Scatter(
    x=df_grouped["Actual Finish"], y=df_grouped["CPI"],
    mode='lines+markers', name="CPI", line=dict(color='green')))
trend_fig.add_trace(go.Scatter(
    x=df_grouped["Actual Finish"], y=df_grouped["SPI"],
    mode='lines+markers', name="SPI", line=dict(color='blue')))
trend_fig.update_layout(
    title="Performance Indices Over Time (CPI & SPI)",
    xaxis_title="Date",
    yaxis_title="Index Value",
    yaxis=dict(range=[0, 2]),
    legend=dict(x=0.01, y=0.99)
)

# ---------------- Save to HTML ----------------
with open("Project_Visualizer_Report.html", "w") as f:
    f.write(gantt_fig.to_html(full_html=False, include_plotlyjs='cdn'))
    f.write(milestone_fig.to_html(full_html=False, include_plotlyjs=False))
    f.write(s_curve_fig.to_html(full_html=False, include_plotlyjs=False))
    f.write(hist_fig.to_html(full_html=False, include_plotlyjs=False))
    f.write(summary_fig.to_html(full_html=False, include_plotlyjs=False))
    f.write(trend_fig.to_html(full_html=False, include_plotlyjs=False))

print("✅ All charts saved to Project_Visualizer_Report.html")
