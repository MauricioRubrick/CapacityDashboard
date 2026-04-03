import math
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

st.set_page_config(page_title="Capacity Model Dashboard", layout="wide")

TECH_TOTAL = 8
UTILIZATION = 0.70
BASE_HOURS = 2000
HOURS_PER_TECH = BASE_HOURS * UTILIZATION
DEFAULT_FLYER_TRAVEL = 200

VISIBLE_FTE_BY_REGION = {
    "Northeast": 1,
    "Southeast": 1,
    "North Central": 1,
    "Midwest": 2,
    "South Central": 1,
    "West": 1,
    "CANADA": 1,
}

TRUCK_COLS = ["Tuggy", "Lowy MC", "Lowy 1171", "Reachy", "Veeny"]

st.title("Capacity Model Dashboard")
st.caption("Regional pressure + national flying technician translation")

uploaded_file = st.file_uploader(
    "Upload BALYO Capacity Workbook",
    type=["xlsx"]
)

if uploaded_file:
    projects = pd.read_excel(uploaded_file, sheet_name="Projects_NA")
    cap = pd.read_excel(uploaded_file, sheet_name="Capacity Model", header=None)

    st.success("Workbook loaded successfully")

    truck_headers = cap.iloc[0, 1:6].tolist()
    expected_vals = cap.iloc[6, 1:6].tolist()
    lost_vals = cap.iloc[7, 1:6].tolist()
    fleet_vals = cap.iloc[2, 1:6].tolist()

    service_factor = {}
    lost_per_robot = []

    for truck, exp, lost, fleet in zip(
        truck_headers, expected_vals, lost_vals, fleet_vals
    ):
        if pd.isna(fleet) or fleet == 0:
            continue

        truck_name = str(truck).strip()
        hrs_per_robot = (float(exp) + float(lost)) / float(fleet)

        service_factor[truck_name] = hrs_per_robot
        lost_per_robot.append(float(lost) / float(fleet))

    avg_hrs_per_robot = sum(service_factor.values()) / len(service_factor)
    avg_flyer_travel = (
        sum(lost_per_robot) / len(lost_per_robot)
        if lost_per_robot else DEFAULT_FLYER_TRAVEL
    )

    flyer_capacity = HOURS_PER_TECH - avg_flyer_travel

    rows = []

    for _, row in projects.iterrows():
        region = str(row.get("Region", "")).strip()

        raw_robots = 0
        weighted_hours = 0

        for truck in TRUCK_COLS:
            qty = row.get(truck, 0)
            qty = 0 if pd.isna(qty) else float(qty)

            raw_robots += qty
            weighted_hours += qty * service_factor.get(truck, 0)

        rows.append({
            "Region": region,
            "Raw Robots": raw_robots,
            "Weighted Hours": weighted_hours,
        })

    df = pd.DataFrame(rows)
    df = df.groupby("Region", as_index=False).sum()

    df["Visible Techs"] = df["Region"].map(
        lambda r: VISIBLE_FTE_BY_REGION.get(r, 0)
    )

    df["Capacity"] = df["Visible Techs"] * HOURS_PER_TECH
    df["Gap"] = (df["Weighted Hours"] - df["Capacity"]).clip(lower=0)

    df["Techs Required"] = (
        df["Weighted Hours"] / HOURS_PER_TECH
    ).round(2)

    df["Hire Need"] = (
        df["Gap"] / HOURS_PER_TECH
    ).apply(lambda x: max(0, math.ceil(x)))

    df["Hours Threshold"] = (
        (df["Visible Techs"] + 1) * HOURS_PER_TECH
    )

    df["Hire Threshold (Robots)"] = (
        df["Hours Threshold"] / avg_hrs_per_robot
    ).round(0)

    df["Covered Demand"] = df[["Weighted Hours", "Capacity"]].min(axis=1)

    df["Over Threshold"] = (
        df["Raw Robots"] - df["Hire Threshold (Robots)"]
    ).clip(lower=0) * avg_hrs_per_robot

    total_weighted = df["Weighted Hours"].sum()
    total_capacity = TECH_TOTAL * HOURS_PER_TECH
    national_gap = max(0, total_weighted - total_capacity)

    national_flying_need = math.ceil(national_gap / flyer_capacity)
    regional_hire_sum = df["Hire Need"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Weighted Demand", f"{total_weighted:,.0f} h")
    c2.metric("Current Capacity", f"{total_capacity:,.0f} h")
    c3.metric("Regional Hire Pressure", int(regional_hire_sum))
    c4.metric("National Flying Tech Need", national_flying_need)

    st.subheader("Demand Coverage, White-Space Risk and Hire Trigger")

    max_hours = max(
        df["Weighted Hours"].max(),
        df["Hours Threshold"].max()
    ) * 1.15

    max_robots = max_hours / avg_hrs_per_robot

    fig = go.Figure()

    fig.add_bar(
        x=df["Region"],
        y=df["Covered Demand"],
        name="Covered Demand",
        marker_color="#6EC1FF"
    )

    fig.add_bar(
        x=df["Region"],
        y=df["Gap"],
        name="Gap",
        marker_color="#FFB347"
    )

    fig.add_bar(
        x=df["Region"],
        y=df["Over Threshold"],
        name="Over Threshold",
        marker_color="#FF4D4D"
    )

    fig.add_scatter(
        x=df["Region"],
        y=df["Hours Threshold"],
        mode="lines+markers",
        name="Hours Threshold",
        line=dict(color="white", width=4, dash="dash"),
        marker=dict(size=10),
        yaxis="y1"
    )

    fig.add_scatter(
        x=df["Region"],
        y=df["Hire Threshold (Robots)"],
        mode="lines+markers",
        name="Robot Threshold",
        line=dict(color="#C084FC", width=4),
        marker=dict(size=10),
        yaxis="y2"
    )

    fig.add_scatter(
        x=df["Region"],
        y=df["Raw Robots"],
        mode="lines+markers",
        name="Robot Qty",
        line=dict(color="#22C55E", width=4),
        marker=dict(size=10),
        yaxis="y2"
    )

    fig.update_layout(
        barmode="stack",
        yaxis=dict(
            title="Weighted Hours",
            range=[0, max_hours]
        ),
        yaxis2=dict(
            title="Robots",
            overlaying="y",
            side="right",
            range=[0, max_robots]
        ),
        title="Demand Coverage, White-Space Risk and Hire Trigger",
        legend=dict(orientation="h", y=1.1)
    )

    st.plotly_chart(fig, use_container_width=True)

    # =========================
    # EXECUTIVE PIVOT TABLE
    # =========================
    st.subheader("Regional KPI Pivot")

    pivot_cols = [
        "Region",
        "Raw Robots",
        "Weighted Hours",
        "Visible Techs",
        "Capacity",
        "Gap",
        "Techs Required",
        "Hire Need",
        "Hire Threshold (Robots)"
    ]

    pivot_df = (
        df[pivot_cols]
        .sort_values("Gap", ascending=False)
        .reset_index(drop=True)
    )

    st.dataframe(
        pivot_df,
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # BULLET SUMMARY BY REGION
    # =========================
    st.subheader("Hiring Justification by Region")

    for _, r in df.sort_values("Gap", ascending=False).iterrows():
        st.markdown(
            f"""
### {r['Region']}
- **{int(r['Raw Robots'])} robots**
- **{int(r['Visible Techs'])} visible techs**
- **weighted demand = {r['Weighted Hours']:,.0f} hrs**
- **capacity = {r['Capacity']:,.0f} hrs**
- **gap = {r['Gap']:,.0f} hrs**
- **Techs Required = {r['Techs Required']}**
- **Hire Need = {int(r['Hire Need'])}**
- **Hire Threshold (Robots) = {int(r['Hire Threshold (Robots)'])}**
"""
        )

    # =========================
    # RECOMMENDATION TO COMEX
    # =========================
    hotspot = df.sort_values("Gap", ascending=False).iloc[0]

    region_name = hotspot["Region"]
    robots_now = int(hotspot["Raw Robots"])
    robot_threshold = int(hotspot["Hire Threshold (Robots)"])
    region_gap = hotspot["Gap"]

    threshold_status = (
        "remains below"
        if robots_now < robot_threshold
        else "has exceeded"
    )

    st.subheader("Recommendation to COMEX")
    st.write(
        f"""
        Current weighted demand is **{total_weighted:,.0f} hrs**
        versus **{total_capacity:,.0f} hrs** of available capacity.

        This creates **{national_gap:,.0f} hrs of national white-space**,
        equivalent to **{national_flying_need} additional flying technicians**.

        **{region_name}** currently carries the highest regional gap
        at **{region_gap:,.0f} hrs** and
        **{threshold_status} its next hiring threshold**
        (**{robots_now} vs {robot_threshold} robots**).
        """
    )

