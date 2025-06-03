import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide")
st.title("Primavera P6 Project Visualizer")

uploaded_file = st.file_uploader("📂 Upload Your Primavera P6 Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Project Data")

    # Convert date columns
    df["Baseline Start"] = pd.to_datetime(df["Baseline Start"])
    df["Baseline Finish"] = pd.to_datetime(df["Baseline Finish"])
    df["Actual Start"] = pd.to_datetime(df["Actual Start"])
    df["Actual Finish"] = pd.to_datetime(df["Actual Finish"])
    
    st.sidebar.title("📊 Filters")
    selected_wbs = st.sidebar.multiselect("Filter by WBS", df["WBS"].unique(), default=df["WBS"].unique())
    df = df[df["WBS"].isin(selected_wbs)]
    selected_area = st.sidebar.multiselect("Filter by Area", df["Area"].unique(), default=df["Area"].unique())
    df = df[df["Area"].isin(selected_area)]
    selected_resource_name = st.sidebar.multiselect("Filter by Resource Name", df["Resource Name"].unique(), default=df["Resource Name"].unique())
    df = df[df["Resource Name"].isin(selected_resource_name)]

    st.subheader("📅 Gantt Chart")
    
    df_sorted = df.sort_values(by="Activity ID")
    
    # Create a new column for color coding
    df["Critical Color"] = df["Critical "].fillna("").apply(lambda x: "critical" if str(x).strip().lower() == "yes" else "Non-Critical")
    
    # Set custom order for y-axis (activity names)
    activity_order = df_sorted["Activity Name"].tolist()

    # Plot using the new color column
    gantt_fig = px.timeline(
        df,
        x_start="Actual Start",
        x_end="Actual Finish",
        y="Activity Name",
        color="Critical Color",  # Use color based on criticality
        color_discrete_map={
        "Critical": "red",
        "Non-Critical": "green"
        },
        category_orders={"Activity Name": activity_order}, 
        hover_data=["Activity ID", "Activity Code", "Planned %", "Actual %", "Remarks"]
    )

    gantt_fig.update_layout(
        yaxis_autorange= "reversed",
        title="Gantt Chart with Critical Path Highlighted",
        height=500
    )
    st.plotly_chart(gantt_fig, use_container_width=True)


    st.subheader("📈 S-Curve (Planned vs. Actual Cost)")
    df_sorted = df.sort_values("Actual Finish")
    df_sorted["Cumulative Planned Cost"] = df_sorted["Budgeted Cost"].cumsum()
    df_sorted["Cumulative Actual Cost"] = df_sorted["Actual Cost"].cumsum()
    s_curve_fig = go.Figure()
    s_curve_fig.add_trace(go.Scatter(x=df_sorted["Actual Finish"], y=df_sorted["Cumulative Planned Cost"],
                                     mode='lines+markers', name='Planned Cost'))
    s_curve_fig.add_trace(go.Scatter(x=df_sorted["Actual Finish"], y=df_sorted["Cumulative Actual Cost"],
                                     mode='lines+markers', name='Actual Cost'))
    s_curve_fig.update_layout(xaxis_title="Date", yaxis_title="Cumulative Cost")
    st.plotly_chart(s_curve_fig, use_container_width=True)

    st.subheader("📊 Man-hours Histogram")
    hist_fig = px.bar(df, x="Activity Name", y=["Budgeted Hours", "Actual Hours"], barmode="group")
    st.plotly_chart(hist_fig, use_container_width=True)

    st.subheader("📌 Earned Value Management")
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

    st.metric("Planned Value (PV)", f"${PV:,.2f}")
    st.metric("Earned Value (EV)", f"${EV:,.2f}")
    st.metric("Actual Cost (AC)", f"${AC:,.2f}")
    st.metric("Cost Variance (CV)", f"${CV:,.2f}")
    st.metric("Schedule Variance (SV)", f"${SV:,.2f}")
    st.metric("CPI", f"{CPI:.2f}" if CPI else "N/A")
    st.metric("SPI", f"{SPI:.2f}" if SPI else "N/A")
    st.metric("Earned Schedule (ES)", f"{ES:.2f}" if ES else "N/A")
    st.metric("Schedule Variance (Time)", f"{Schedule_Variance_Time:.2f}" if Schedule_Variance_Time else "N/A")

    st.subheader("📈 Performance Indices Over Time")
    df_grouped = df.groupby("Actual Finish").agg({"EV": "sum", "PV": "sum", "AC": "sum"}).sort_index().reset_index()
    df_grouped["CPI"] = df_grouped["EV"] / df_grouped["AC"]
    df_grouped["SPI"] = df_grouped["EV"] / df_grouped["PV"]
    trend_fig = go.Figure()
    trend_fig.add_trace(go.Scatter(x=df_grouped["Actual Finish"], y=df_grouped["CPI"],
                                   mode='lines+markers', name="CPI", line=dict(color='green')))
    trend_fig.add_trace(go.Scatter(x=df_grouped["Actual Finish"], y=df_grouped["SPI"],
                                   mode='lines+markers', name="SPI", line=dict(color='blue')))
    trend_fig.update_layout(xaxis_title="Date", yaxis_title="Index Value", yaxis=dict(range=[0, 2]))
    st.plotly_chart(trend_fig, use_container_width=True)

else:
    st.info("Please upload a Primavera P6 Excel file to get started.")
