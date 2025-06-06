"""Office Attendance Analyzer – streamlined & robust
=====================================================
A Streamlit web‑app that ingests a standard attendance export (e.g. SAP
SuccessFactors) and produces per‑person / team metrics for both ISO weeks and
for the current calendar month.

Rev 4 (finalized)
-----------------
* Handles holidays and absences accurately in week-level breakdowns.
* Displays % values in UI.
* Red text for values <60%.
* Compatible with all pandas versions.
* Suppresses unnecessary trailing decimal zeros for cleaner display.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Final, Iterable

import pandas as pd
import streamlit as st
from pandas.tseries.offsets import MonthBegin

try:
    import holidays  # type: ignore
except ModuleNotFoundError as exc:  # pragma: no cover
    raise ModuleNotFoundError(
        "pip install holidays  – required for public‑holiday look‑ups"
    ) from exc

###############################################################################
# CONFIGURATION
###############################################################################
COUNTRY_HOLIDAYS: Final[str] = "EE"  # ISO‑3166 alpha‑2 – set yours here
DAILY_EXPECTED_HOURS: Final[float] = 8.0
LOW_PCT_THRESHOLD: Final[float] = 0.60  # red‑text threshold for % columns

ABSENCE_KEYWORDS: Final[set[str]] = {
    "vacation",
    "annual leave",
    "long service award",
    "military leave",
    "sick leave",
}

###############################################################################
# HELPER FUNCTIONS
###############################################################################

def _business_days(start: date, end: date, holiday_set: set[date]) -> int:
    rng = pd.bdate_range(start, end, freq="C", weekmask="Mon Tue Wed Thu Fri")
    return sum(d.date() not in holiday_set for d in rng)

def working_days_month(year: int, month: int, holiday_set: set[date]) -> int:
    first = date(year, month, 1)
    last = (datetime(year, month, 1) + MonthBegin(1) - timedelta(days=1)).date()
    return _business_days(first, last, holiday_set)

def working_days_iso_week(year: int, iso_week: int, holiday_set: set[date]) -> int:
    monday = date.fromisocalendar(year, iso_week, 1)
    friday = monday + timedelta(days=4)
    return _business_days(monday, friday, holiday_set)

def is_absence(text: str | int | float | None) -> bool:
    return any(k in str(text or "").lower() for k in ABSENCE_KEYWORDS)

###############################################################################
# DATA LOADERS (cached)
###############################################################################

@st.cache_data(show_spinner="Parsing Excel …")
def load_xlsx(buffer: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(buffer))

@st.cache_data(show_spinner="Crunching numbers …")
def process_attendance(raw: pd.DataFrame):
    df = raw.copy()

    df["Attendance date"] = pd.to_datetime(df["Attendance date"], format="%d.%m.%Y", errors="coerce")
    df["HoursWorked"] = pd.to_numeric(df["Total time worked decimal value"], errors="coerce").fillna(0.0)
    df["Present"] = df["HoursWorked"] > 0
    df["Vacation"] = df["Event"].map(is_absence)

    years = df["Attendance date"].dt.year.dropna().unique().astype(int)
    hols = {
        d for y in years for d in holidays.country_holidays(COUNTRY_HOLIDAYS, years=[int(y)]).keys()
    }

    df = df[df["Attendance date"].dt.dayofweek < 5]
    df = df[~df["Attendance date"].dt.date.isin(hols)]

    latest = df["Attendance date"].max()
    year, month = latest.year, latest.month
    workdays_month = working_days_month(year, month, hols)
    month_label = latest.strftime("%B %Y")

    df_month = df[(df["Attendance date"].dt.year == year) & (df["Attendance date"].dt.month == month)]

    vac_days_month = (
        df_month[df_month["Vacation"]]
        .groupby("Employee name")["Attendance date"]
        .nunique()
        .rename("VacationDays")
    )

    person_month = (
        df_month.groupby("Employee name")
        .agg(DaysInOffice=("Present", "sum"), ActualHours=("HoursWorked", "sum"))
        .join(vac_days_month, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
    )
    person_month["ExpectedDays"] = workdays_month - person_month["VacationDays"]
    person_month["ExpectedHours"] = person_month["ExpectedDays"] * DAILY_EXPECTED_HOURS
    person_month["PctWorkingDays"] = pd.to_numeric(
        person_month["DaysInOffice"] / person_month["ExpectedDays"].replace(0, pd.NA), errors="coerce"
    ).round(2)
    person_month["PctHours"] = pd.to_numeric(
        person_month["ActualHours"] / person_month["ExpectedHours"].replace(0, pd.NA), errors="coerce"
    ).round(2)

    team_size_month = df_month["Employee name"].nunique()
    vac_pd_month = vac_days_month.sum()
    exp_pd_month = workdays_month * team_size_month - vac_pd_month

    summary_month = pd.DataFrame({
        "Month": [month_label],
        "Working Days": [workdays_month],
        "Team Size": [team_size_month],
        "Vacation Person‑Days": [vac_pd_month],
        "Team Presence %": [
            (df_month["Present"].sum() / exp_pd_month if exp_pd_month else 0.0)
        ],
        "Team Hours %": [
            (df_month["HoursWorked"].sum() / (exp_pd_month * DAILY_EXPECTED_HOURS) if exp_pd_month else 0.0)
        ],
    })

    summary_month["Team Presence %"] = pd.to_numeric(summary_month["Team Presence %"], errors="coerce").fillna(0.0).round(2)
    summary_month["Team Hours %"] = pd.to_numeric(summary_month["Team Hours %"], errors="coerce").fillna(0.0).round(2)

    df["ISOYear"] = df["Attendance date"].dt.isocalendar().year.astype(int)
    df["ISOWeek"] = df["Attendance date"].dt.isocalendar().week.astype(int)

    base_weeks = (
        df[["ISOYear", "ISOWeek"]]
        .drop_duplicates()
        .assign(
            WorkingDays=lambda x: x.apply(
                lambda r: working_days_iso_week(r.ISOYear, r.ISOWeek, hols), axis=1
            )
        )
        .astype({"WorkingDays": "int8"})
        .assign(ExpectedHoursWeek=lambda x: x.WorkingDays * DAILY_EXPECTED_HOURS)
    )

    vac_days_week = (
        df[df["Vacation"]]
        .groupby(["ISOYear", "ISOWeek", "Employee name"])["Attendance date"]
        .nunique()
        .rename("VacationDays")
    )

    person_week = (
        df.groupby(["ISOYear", "ISOWeek", "Employee name"])
        .agg(DaysInOffice=("Present", "sum"), ActualHours=("HoursWorked", "sum"))
        .join(vac_days_week, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
        .merge(base_weeks, on=["ISOYear", "ISOWeek"])
    )
    person_week["ExpectedDays"] = person_week["WorkingDays"] - person_week["VacationDays"]
    person_week["ExpectedHours"] = person_week["ExpectedDays"] * DAILY_EXPECTED_HOURS
    person_week["Year‑Week"] = person_week["ISOYear"].astype(str) + "‑W" + person_week["ISOWeek"].astype(str).str.zfill(2)
    person_week["PctWorkingDays"] = pd.to_numeric(
        person_week["DaysInOffice"] / person_week["ExpectedDays"].replace(0, pd.NA), errors="coerce"
    ).round(2)
    person_week["PctHours"] = pd.to_numeric(
        person_week["ActualHours"] / person_week["ExpectedHours"].replace(0, pd.NA), errors="coerce"
    ).round(2)

    vac_pd_week = vac_days_week.groupby(["ISOYear", "ISOWeek"]).sum().rename("VacPD")

    team_week = (
        df.groupby(["ISOYear", "ISOWeek"])
        .agg(PersonDays=("Present", "sum"), ActualTeamHours=("HoursWorked", "sum"))
        .reset_index()
        .merge(base_weeks, on=["ISOYear", "ISOWeek"])
        .join(vac_pd_week, on=["ISOYear", "ISOWeek"])
        .fillna({"VacPD": 0})
    )
    team_size_all = df["Employee name"].nunique()
    team_week["ExpectedPersonDays"] = team_week["WorkingDays"] * team_size_all - team_week["VacPD"]
    team_week["ExpectedTeamHours"] = team_week["ExpectedPersonDays"] * DAILY_EXPECTED_HOURS
    team_week["Year‑Week"] = team_week["ISOYear"].astype(str) + "‑W" + team_week["ISOWeek"].astype(str).str.zfill(2)
    team_week["TeamPresence%"] = pd.to_numeric(
        team_week.PersonDays / team_week.ExpectedPersonDays.replace(0, pd.NA), errors="coerce"
    ).round(2)
    team_week["TeamHours%"] = pd.to_numeric(
        team_week.ActualTeamHours / team_week.ExpectedTeamHours.replace(0, pd.NA), errors="coerce"
    ).round(2)

    def smart_format(val):
        if pd.isna(val):
            return ""
        elif isinstance(val, float):
            return f"{val:.2f}".rstrip("0").rstrip(".")
        return val

    for df_ in (person_month, person_week, team_week, summary_month):
        for col in df_.select_dtypes(include="number").columns:
            df_[col] = df_[col].apply(smart_format)

    return summary_month, person_month, person_week, team_week

###############################################################################
# PRESENTATION HELPERS
###############################################################################

def style_pct(df: pd.DataFrame, cols: Iterable[str]) -> pd.Styler:
    formatter = {c: "{:.0%}" for c in cols}

    def red(val: float | pd.NA):  # noqa: ANN001
        if pd.isna(val):
            return ""
        return "color:red;" if val < LOW_PCT_THRESHOLD else ""

    return df.style.hide(axis="index").format(formatter).applymap(red, subset=cols)

###############################################################################
# STREAMLIT APP
###############################################################################

def main():  # pragma: no cover
    st.set_page_config(page_title="Attendance Analyzer", layout="wide")
    st.title("📊 Office Attendance Analyzer")

    uploaded = st.file_uploader("Upload attendance report (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    if uploaded is None:
        st.info("👆 Drop or select an .xlsx file to begin")
        st.stop()

    raw_df = load_xlsx(uploaded.getvalue())

    try:
        summary_m, person_m, person_w, team_w = process_attendance(raw_df)
    except Exception as exc:  # pragma: no cover
        st.exception(exc)
        st.stop()

    tab_summary, tab_debug = st.tabs(["📈 Summary", "🪵 Raw Debug"])

    with tab_summary:
        st.subheader("Monthly Snapshot")
        st.dataframe(style_pct(summary_m, ["Team Presence %", "Team Hours %"]), use_container_width=True)

        col_month, col_week = st.columns(2)
        with col_month:
            st.subheader("Per‑Person (Month)")
            st.dataframe(style_pct(person_m, ["PctWorkingDays", "PctHours"]), use_container_width=True)
        with col_week:
            st.subheader("Per‑Person (Week)")
            st.dataframe(style_pct(person_w, ["PctWorkingDays", "PctHours"]), use_container_width=True)

        st.subheader("Team (Week)")
        st.dataframe(style_pct(team_w, ["TeamPresence%", "TeamHours%"]), use_container_width=True)

        st.download_button(
            label="Download Monthly Summary (CSV)",
            data=summary_m.to_csv(index=False).encode(),
            file_name="attendance_summary.csv",
            mime="text/csv",
        )

    with tab_debug:
        st.write("### Raw data (first 100 rows after initial parsing)")
        st.dataframe(raw_df.head(100), use_container_width=True)

if __name__ == "__main__":
    main()