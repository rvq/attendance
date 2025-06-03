import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO
import holidays  # Estonia national holidays

###############################################################################
# CONFIG & CONSTANTS
###############################################################################

DAILY_EXPECTED_HOURS = 8.0  # change if your standard day is different

ee_holidays = holidays.EE()  # all‑year Estonian national & public holidays

###############################################################################
# Helper functions
###############################################################################

def working_days_in_month(year: int, month: int) -> int:
    """Return the count of Monday‑Friday *non‑holiday* days in a given month."""
    cal = calendar.Calendar()
    return sum(
        1
        for day, dow in cal.itermonthdays2(year, month)
        if day and dow < 5 and datetime(year, month, day).date() not in ee_holidays
    )

###############################################################################
# Core processor
###############################################################################

def process_attendance(df: pd.DataFrame):
    """Compute attendance & hours metrics. Expects columns:
    - Employee name
    - Attendance date (dd.mm.yyyy)
    - Time in (can be blank)
    - Total time worked decimal value (hours as float)"""

    df = df.copy()
    df["Attendance date"] = pd.to_datetime(df["Attendance date"], format="%d.%m.%Y")

    # Actual hours (decimal); fallback to 0 for NaNs / blanks
    df["HoursWorked"] = pd.to_numeric(
        df["Total time worked decimal value"], errors="coerce"
    ).fillna(0.0)

    df["Present"] = df["HoursWorked"] > 0  # present if >0 hours

    # Filter out weekends & national holidays
    df = df[df["Attendance date"].dt.dayofweek < 5]
    df = df[~df["Attendance date"].dt.date.isin(ee_holidays)]

    # Determine month context (latest month in sheet)
    latest_date = df["Attendance date"].max()
    year, month = latest_date.year, latest_date.month
    ym_key = pd.Period(datetime(year, month, 1), "M")

    # ---------------------------------------------------------------
    # Working‑day & expected‑hour helpers
    # ---------------------------------------------------------------
    working_days_month = working_days_in_month(year, month)
    expected_hours_month_per_person = working_days_month * DAILY_EXPECTED_HOURS

    # Mask rows of that month
    month_mask = (
        (df["Attendance date"].dt.year == year)
        & (df["Attendance date"].dt.month == month)
    )
    df_month = df[month_mask]

    # ===============================================================
    # Per‑Person · Month
    # ===============================================================
    person_month = (
        df_month.groupby("Employee name")["Present", "HoursWorked"].agg(
            DaysInOffice=("Present", "sum"),
            ActualHours=("HoursWorked", "sum"),
        )
    ).reset_index()

    person_month["ExpectedHours"] = expected_hours_month_per_person
    person_month["PctOfWorkingDays"] = (
        person_month["DaysInOffice"] / working_days_month
    ).round(2)
    person_month["PctOfHours"] = (
        person_month["ActualHours"] / person_month["ExpectedHours"]
    ).round(2)

    # ===============================================================
    # Per‑Person · Week
    # ===============================================================
    df["ISOWeek"] = df["Attendance date"].dt.isocalendar().week.astype(int)

    week_working_days = (
        df.groupby("ISOWeek")["Attendance date"].nunique()
        .rename("WorkingDays")
        .reset_index()
    )
    week_working_days["ExpectedHours"] = (
        week_working_days["WorkingDays"] * DAILY_EXPECTED_HOURS
    )

    person_week = (
        df.groupby(["ISOWeek", "Employee name"])["Present", "HoursWorked"].agg(
            DaysInOffice=("Present", "sum"),
            ActualHours=("HoursWorked", "sum"),
        )
        .reset_index()
        .merge(week_working_days, on="ISOWeek")
    )
    person_week["ExpectedHours"] = person_week["ExpectedHours"]  # merge column
    person_week["PctOfWorkingDays"] = (
        person_week["DaysInOffice"] / person_week["WorkingDays"]
    ).round(2)
    person_week["PctOfHours"] = (
        person_week["ActualHours"] / person_week["ExpectedHours"]
    ).round(2)

    # ===============================================================
    # Team metrics
    # ===============================================================
    team_size = df["Employee name"].nunique()

    # ---- Weekly team hours ----
    team_week = (
        df.groupby("ISOWeek")["Present", "HoursWorked"].agg(
            PersonDays=("Present", "sum"),
            ActualTeamHours=("HoursWorked", "sum"),
        )
        .reset_index()
        .merge(week_working_days, on="ISOWeek")
    )
    team_week["ExpectedTeamHours"] = (
        team_week["WorkingDays"] * DAILY_EXPECTED_HOURS * team_size
    )
    team_week["TeamPresencePct"] = (
        team_week["PersonDays"] / (team_week["WorkingDays"] * team_size)
    ).round(2)
    team_week["TeamHoursPct"] = (
        team_week["ActualTeamHours"] / team_week["ExpectedTeamHours"]
    ).round(2)

    # ---- Monthly team hours ----
    actual_team_hours_month = df_month["HoursWorked"].sum()
    expected_team_hours_month = (
        expected_hours_month_per_person * team_size
    )
    team_month_df = pd.DataFrame(
        {
            "YearMonth": [str(ym_key)],
            "PersonDays": [df_month["Present"].sum()],
            "WorkingDays": [working_days_month],
            "TeamSize": [team_size],
            "TeamPresencePct": [round(
                df_month["Present"].sum() / (working_days_month * team_size), 2)
            ],
            "ActualTeamHours": [actual_team_hours_month],
            "ExpectedTeamHours": [expected_team_hours_month],
            "TeamHoursPct": [round(
                actual_team_hours_month / expected_team_hours_month, 2)
            ],
        }
    )

    # ===============================================================
    # Summary
    # ===============================================================
    summary_df = pd.DataFrame(
        {
            "Month": [ym_key.strftime("%B %Y")],
            "Working Days": [working_days_month],
            "Team Size": [team_size],
            "Team Presence %": [team_month_df := team_month_df if 'team_month_df' in locals() else None],  # placeholder to maintain existing col order
            "Actual Team Hours": [actual_team_hours_month],
            "Expected Team Hours": [expected_team_hours_month],
            "Team Hours %": [round(actual_team_hours_month / expected_team_hours_month, 2)],
        }
    )
    # Remove placeholder column if accidentally None
    if summary_df["Team Presence %"].isnull().all():
        summary_df.drop(columns=["Team Presence %"], inplace=True)
        summary_df.insert(3, "Team Presence %", team_month_df["TeamPresencePct"].iloc[0])

    return (
        summary_df,
        person_month,
        person_week,
        team_week,
        team_month_df,
    )

###############################################################################
# Streamlit UI
###############################################################################

def main():
    st.set_page_config(page_title="Office Attendance Analyzer", layout="wide")
    st.title("📊 Office Attendance Analyzer")

    uploaded_file = st.file_uploader(
        "Upload attendance report (.xlsx)", type=["xlsx"]
    )
    if uploaded_file is None:
        st.info("👆 Drop a file here or click to select")
        st.stop()

    # Read the Excel into DataFrame using BytesIO buffer
    df = pd.read_excel(BytesIO(uploaded_file.read()))

    try:
        (
            summary_df,
            person_month,
            person_week,
            team_week,
            team_month_df,
        ) = process_attendance(df)
    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")
        st.stop()

    st.subheader("Monthly Summary")
    st.dataframe(summary_df, hide_index=True)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Per‑Person (Month)")
        st.dataframe(person_month, hide_index=True)

    with col2:
        st.subheader("Team Presence & Hours (Month)")
        st.dataframe(team_month_df, hide_index=True)

    st.subheader("Per‑Person (Week)")
    st.dataframe(person_week, hide_index=True)

    st.subheader("Team Presence & Hours (Week)")
    st.dataframe(team_week, hide_index=True)

    # Allow user to download summary
    csv = summary_df.to...
