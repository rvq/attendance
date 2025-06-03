import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO

###############################################################################
# Helper functions
###############################################################################

def working_days_in_month(year: int, month: int) -> int:
    """Return number of Monday‑Friday days in a calendar month."""
    cal = calendar.Calendar()
    return sum(1 for day, dow in cal.itermonthdays2(year, month) if day and dow < 5)


def process_attendance(df: pd.DataFrame):
    """Compute the various attendance metrics and return the five DataFrames
    we want to show in the UI.
    Expects columns:  Employee name | Attendance date | Time in"""

    # Normalise / parse
    df = df.copy()
    df["Attendance date"] = pd.to_datetime(df["Attendance date"], format="%d.%m.%Y")
    df["Present"] = df["Time in"].notna()
    df = df[df["Attendance date"].dt.dayofweek < 5]  # Mon‑Fri only

    # Core keys
    latest_date = df["Attendance date"].max()
    year, month = latest_date.year, latest_date.month
    ym_key = pd.Period(datetime(year, month, 1), "M")

    # ---------- Monthly working‑days ----------------------------------------
    working_days_month = working_days_in_month(year, month)

    # ---------- Per‑Person ‑ whole month -----------------------------------
    person_month = (df[(df["Attendance date"].dt.year == year) & (df["Attendance date"].dt.month == month)]
                    .groupby("Employee name")["Present"].sum()
                    .reset_index(name="DaysInOffice"))
    person_month["PctOfWorkingDays"] = (person_month["DaysInOffice"] / working_days_month).round(2)

    # ---------- Per‑Person ‑ by week --------------------------------------
    df["ISOWeek"] = df["Attendance date"].dt.isocalendar().week.astype(int)
    week_working_days = (df.groupby("ISOWeek")["Attendance date"].nunique()
                         .rename("WorkingDays").reset_index())

    person_week = (df.groupby(["ISOWeek", "Employee name"])["Present"].sum()
                   .reset_index(name="DaysInOffice")
                   .merge(week_working_days, on="ISOWeek"))
    person_week["PctOfWorkingDays"] = (person_week["DaysInOffice"] / person_week["WorkingDays"]).round(2)

    # ---------- Team presence ‑ weekly -------------------------------------
    team_size = df["Employee name"].nunique()
    team_week = (df.groupby("ISOWeek")["Present"].sum()
                 .reset_index(name="PersonDays")
                 .merge(week_working_days, on="ISOWeek"))
    team_week["TeamPresencePct"] = (team_week["PersonDays"] /
                                     (team_week["WorkingDays"] * team_size)).round(2)

    # ---------- Team presence ‑ monthly ------------------------------------
    team_month_persondays = df[(df["Attendance date"].dt.year == year) & (df["Attendance date"].dt.month == month)]["Present"].sum()
    team_presence_month_pct = round(team_month_persondays / (working_days_month * team_size), 2)

    summary_df = pd.DataFrame({
        "Month": [ym_key.strftime("%B %Y")],
        "Working Days": [working_days_month],
        "Team Size": [team_size],
        "Team Presence %": [team_presence_month_pct]
    })

    team_month_df = pd.DataFrame({
        "YearMonth": [str(ym_key)],
        "PersonDays": [team_month_persondays],
        "TeamSize": [team_size],
        "WorkingDays": [working_days_month],
        "TeamPresencePct": [team_presence_month_pct]
    })

    return summary_df, person_month, person_week, team_week, team_month_df

###############################################################################
# Streamlit UI
###############################################################################

def main():
    st.set_page_config(page_title="Office Attendance Analyzer", layout="wide")
    st.title("📊 Office Attendance Analyzer")

    uploaded_file = st.file_uploader("Upload attendance report (.xlsx)", type=["xlsx"])
    if uploaded_file is None:
        st.info("👆 Drop a file here or click to select")
        st.stop()

    # Read the Excel into DataFrame using BytesIO buffer
    df = pd.read_excel(BytesIO(uploaded_file.read()))

    try:
        summary_df, person_month, person_week, team_week, team_month_df = process_attendance(df)
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
        st.subheader("Team Presence (Month)")
        st.dataframe(team_month_df, hide_index=True)

    st.subheader("Per‑Person (Week)")
    st.dataframe(person_week, hide_index=True)

    st.subheader("Team Presence (Week)")
    st.dataframe(team_week, hide_index=True)

    # Allow user to download summary
    csv = summary_df.to_csv(index=False).encode()
    st.download_button("Download Monthly Summary (CSV)", csv, file_name="attendance_summary.csv", mime="text/csv")


if __name__ == "__main__":
    main()
