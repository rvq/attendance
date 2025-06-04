import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, timedelta, date
from io import BytesIO
import holidays  # Estonia national holidays

###############################################################################
# CONFIG
###############################################################################

DAILY_EXPECTED_HOURS = 8.0  # tweak if a full work‑day ≠ 8 h
EE_HOLIDAYS = holidays.EE()  # public holidays for every year

# Any Event values that should count as paid leave / absence.
ABSENCE_KEYWORDS = {
    "vacation",
    "annual leave",
    "long service award",
    "military leave"
}

# Threshold below which percentages are highlighted in red
THRESHOLD_LOW_PCT = 0.60

###############################################################################
# Helper functions
###############################################################################

def working_days_in_month(year: int, month: int) -> int:
    """Return count of Mon–Fri days in the month minus public holidays."""
    cal = calendar.Calendar()
    return sum(
        1
        for day, dow in cal.itermonthdays2(year, month)
        if day and dow < 5 and datetime(year, month, day).date() not in EE_HOLIDAYS
    )


def working_days_in_iso_week(year: int, iso_week: int) -> int:
    """Return Mon–Fri business-day count in the ISO week, minus Estonian public holidays."""
    monday = date.fromisocalendar(year, iso_week, 1)
    return sum(
        1
        for i in range(5)  # Monday‑Friday
        if (monday + timedelta(days=i)) not in EE_HOLIDAYS
    )

###############################################################################
# Core processor
###############################################################################

def process_attendance(df: pd.DataFrame):
    """Compute attendance & hour metrics from the raw DataFrame.

    Required columns (exact names):
      • Employee name
      • Attendance date                 (dd.mm.yyyy)
      • Total time worked decimal value (float hours, blanks ok)
      • Event                            (absence reason)
    """

    df = df.copy()
    df["Attendance date"] = pd.to_datetime(df["Attendance date"], format="%d.%m.%Y")

    # Hours & presence flag --------------------------------------------------
    df["HoursWorked"] = pd.to_numeric(df["Total time worked decimal value"], errors="coerce").fillna(0.0)
    df["Present"] = df["HoursWorked"] > 0

    # Absence detection ------------------------------------------------------
    df["EventNorm"] = df["Event"].fillna("").str.strip().str.lower()
    df["Vacation"] = df["EventNorm"].isin(ABSENCE_KEYWORDS)

    # Exclude weekends and public holidays (absences remain for counting) ----
    df = df[df["Attendance date"].dt.dayofweek < 5]
    df = df[~df["Attendance date"].dt.date.isin(EE_HOLIDAYS)]

    # ----------------------------------------------------------------------
    # MONTHLY CALCULATIONS (unchanged)
    # ----------------------------------------------------------------------
    latest = df["Attendance date"].max()
    year, month = latest.year, latest.month
    ym_period = pd.Period(datetime(year, month, 1), "M")

    working_days_month = working_days_in_month(year, month)

    month_mask = (df["Attendance date"].dt.year == year) & (df["Attendance date"].dt.month == month)
    df_month = df[month_mask]

    vac_days_month = df_month.groupby("Employee name")["Vacation"].sum().rename("VacationDays")

    person_month = (
        df_month.groupby("Employee name")
        .agg(DaysInOffice=("Present", "sum"), ActualHours=("HoursWorked", "sum"))
        .join(vac_days_month, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
    )
    person_month["ExpectedDays"] = working_days_month - person_month["VacationDays"]
    person_month["ExpectedHours"] = person_month["ExpectedDays"] * DAILY_EXPECTED_HOURS
    person_month["PctOfWorkingDays"] = (person_month["DaysInOffice"] / person_month["ExpectedDays"].replace(0, pd.NA)).round(2)
    person_month["PctOfHours"] = (person_month["ActualHours"] / person_month["ExpectedHours"].replace(0, pd.NA)).round(2)

    # ----------------------------------------------------------------------
    # WEEKLY CALCULATIONS (refactored)
    # ----------------------------------------------------------------------
    df["ISOWeek"] = df["Attendance date"].dt.isocalendar().week.astype(int)
    df["ISOYear"] = df["Attendance date"].dt.isocalendar().year.astype(int)

    # Build calendar‑based working‑day reference for every ISO year/week present in data
    unique_weeks = df[["ISOYear", "ISOWeek"]].drop_duplicates()
    unique_weeks["WorkingDays"] = unique_weeks.apply(
        lambda r: working_days_in_iso_week(int(r.ISOYear), int(r.ISOWeek)), axis=1
    )
    unique_weeks["ExpectedHoursWeek"] = unique_weeks["WorkingDays"] * DAILY_EXPECTED_HOURS

    # Per‑person vacation days per week ------------------------------------
    vac_days_week = (
        df.groupby(["ISOYear", "ISOWeek", "Employee name"])["Vacation"].sum().rename("VacationDays")
    )

    # PER‑PERSON · WEEK -----------------------------------------------------
    person_week = (
        df.groupby(["ISOYear", "ISOWeek", "Employee name"])
        .agg(DaysInOffice=("Present", "sum"), ActualHours=("HoursWorked", "sum"))
        .join(vac_days_week, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
        .merge(unique_weeks, on=["ISOYear", "ISOWeek"])
    )
    person_week["ExpectedDays"] = person_week["WorkingDays"] - person_week["VacationDays"]
    person_week["ExpectedHours"] = person_week["ExpectedDays"] * DAILY_EXPECTED_HOURS
    person_week["PctOfWorkingDays"] = (person_week["DaysInOffice"] / person_week["ExpectedDays"].replace(0, pd.NA)).round(2)
    person_week["PctOfHours"] = (person_week["ActualHours"] / person_week["ExpectedHours"].replace(0, pd.NA)).round(2)
    # Nice label for UI
    person_week["YearWeek"] = person_week["ISOYear"].astype(str) + "‑W" + person_week["ISOWeek"].astype(str).str.zfill(2)

    # TEAM‑LEVEL · WEEK -----------------------------------------------------
    team_size = df["Employee name"].nunique()

    vac_persondays_week = (
        df.groupby(["ISOYear", "ISOWeek"])["Vacation"].sum().rename("VacationPersonDays")
    )

    team_week = (
        df.groupby(["ISOYear", "ISOWeek"])
        .agg(PersonDays=("Present", "sum"), ActualTeamHours=("HoursWorked", "sum"))
        .reset_index()
        .merge(unique_weeks, on=["ISOYear", "ISOWeek"])
        .join(vac_persondays_week, on=["ISOYear", "ISOWeek"])
        .fillna({"VacationPersonDays": 0})
    )
    team_week["ExpectedPersonDays"] = team_week["WorkingDays"] * team_size - team_week["VacationPersonDays"]
    team_week["ExpectedTeamHours"] = team_week["ExpectedPersonDays"] * DAILY_EXPECTED_HOURS
    team_week["TeamPresencePct"] = (team_week["PersonDays"] / team_week["ExpectedPersonDays"].replace(0, pd.NA)).round(2)
    team_week["TeamHoursPct"] = (team_week["ActualTeamHours"] / team_week["ExpectedTeamHours"].replace(0, pd.NA)).round(2)
    # Label
    team_week["YearWeek"] = team_week["ISOYear"].astype(str) + "‑W" + team_week["ISOWeek"].astype(str).str.zfill(2)

    # ----------------------------------------------------------------------
    # TEAM‑LEVEL · MONTH (unchanged)
    # ----------------------------------------------------------------------
    team_size_month = df_month["Employee name"].nunique()
    total_vac_persondays_month = df_month["Vacation"].sum()
    actual_team_hours_month = df_month["HoursWorked"].sum()
    expected_persondays_month = working_days_month * team_size_month - total_vac_persondays_month
    expected_team_hours_month = expected_persondays_month * DAILY_EXPECTED_HOURS
    team_presence_pct_month = (
        df_month["Present"].sum() / expected_persondays_month if expected_persondays_month else pd.NA
    )

    team_month_df = pd.DataFrame({
        "YearMonth": [str(ym_period)],
        "PersonDays": [df_month["Present"].sum()],
        "ExpectedPersonDays": [expected_persondays_month],
        "TeamSize": [team_size_month],
        "VacationPersonDays": [total_vac_persondays_month],
        "TeamPresencePct": [round(team_presence_pct_month, 2) if pd.notna(team_presence_pct_month) else pd.NA],
        "ActualTeamHours": [actual_team_hours_month],
        "ExpectedTeamHours": [expected_team_hours_month],
        "TeamHoursPct": [
            round(actual_team_hours_month / expected_team_hours_month, 2) if expected_team_hours_month else pd.NA
        ],
    })

    summary_df = pd.DataFrame({
        "Month": [ym_period.strftime("%B %Y")],
        "Working Days": [working_days_month],
        "Team Size": [team_size_month],
        "Vacation Person‑Days": [total_vac_persondays_month],
        "Team Presence %": [team_month_df["TeamPresencePct"].iloc[0]],
        "Team Hours %": [team_month_df["TeamHoursPct"].iloc[0]],
    })

    return summary_df, person_month, person_week, team_week, team_month_df

###############################################################################
# Styling helper
###############################################################################

def style_attendance(df: pd.DataFrame, percent_cols: list[str]):
    """Return a Styler with percentage formatting and red text < 60%."""
    formatter = {col: "{:.0%}" for col in percent_cols}

    def highlight_low(v):
        if pd.notna(v) and v < THRESHOLD_LOW_PCT:
            return "color: red;"
        return ""

    return (
        df.style
        .format(formatter)
        .applymap(highlight_low, subset=percent_cols)
        .hide(axis="index")
    )

###############################################################################
# Streamlit UI
###############################################################################

def main():
    st.set_page_config(page_title="📊 Office Attendance Analyzer", layout="wide")
    st.title("📊 Office Attendance Analyzer")

    uploaded_file = st.file_uploader("Upload attendance report (.xlsx)", type=["xlsx"])
    if uploaded_file is None:
        st.info("👆 Drop a file here or click to select")
        st.stop()

    df = pd.read_excel(BytesIO(uploaded_file.read()))

    if st.checkbox("Show zero‑hour Event counts"):
        st.write(
            df[df["Total time worked decimal value"].fillna(0) == 0]["Event"].value_counts(dropna=False).head(20)
        )

    try:
        summary_df, person_month, person_week, team_week, team_month_df = process_attendance(df)
    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")
        st.stop()


    # ---------------------------- Outputs -----------------------------------
    st.subheader("Monthly Summary")
    st.dataframe(style_attendance(summary_df, ["Team Presence %", "Team Hours %"]), use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Per-Person (Month)")
        st.dataframe(
            style_attendance(person_month, ["PctOfWorkingDays", "PctOfHours"]),
            use_container_width=True,
        )
    with col2:
        st.subheader("Team Presence & Hours (Month)")
        st.dataframe(
            style_attendance(team_month_df, ["TeamPresencePct", "TeamHoursPct"]),
            use_container_width=True,
        )

    st.subheader("Per-Person (Week)")
    st.dataframe(
        style_attendance(person_week, ["PctOfWorkingDays", "PctOfHours"]),
        use_container_width=True,
    )

    st.subheader("Team Presence & Hours (Week)")
    st.dataframe(
        style_attendance(team_week, ["TeamPresencePct", "TeamHoursPct"]),
        use_container_width=True,
    )

    # ---------------------------- Download button ----------------------------
    csv_bytes = summary_df.to_csv(index=False).encode()
    st.download_button(
        label="Download Monthly Summary (CSV)",
        data=csv_bytes,
        file_name="attendance_summary.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
