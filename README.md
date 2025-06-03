import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Load the Excel file
file_path = "Sample_Primavera_Data.xlsx"  # Update this if your file name is different
df = pd.read_excel(file_path, sheet_name="Project Data")

# Convert date columns to datetime
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
milestones = df.iloc[::5].copy()  # take every 5th row as sample milestone
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

# ---------------- Save all charts to one HTML ----------------
with open("Project_Visualizer_Report.html", "w") as f:
    f.write(gantt_fig.to_html(full_html=False, include_plotlyjs='cdn'))
    f.write(milestone_fig.to_html(full_html=False, include_plotlyjs=False))
    f.write(s_curve_fig.to_html(full_html=False, include_plotlyjs=False))
    f.write(hist_fig.to_html(full_html=False, include_plotlyjs=False))

print("✅ Report saved as: Project_Visualizer_Report.html")
