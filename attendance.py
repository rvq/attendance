"""Office Attendance Analyzer – streamlined & robust
====================================================
A one‑page Streamlit web‑app that ingests a standard *SAP SuccessFactors* (or
similar) attendance export and produces per‑person / team metrics for both ISO
weeks and the current calendar month.

Key improvements over earlier drafts
------------------------------------
* **End‑to‑end typing safety** – every column is cast exactly once; math then
  happens exclusively on numeric dtypes.
* **Calendar‑first logic** – working‑day counts come from the calendar, *never*
  from whatever rows happen to be in the import.
* **Substring absence matching** – covers “Vacation – Summer Trip” etc.
* **@st.cache_data** wrappers so large files re‑parse instantly on rerun.
* **Cleaner UI** – two tabs ("Summary" & "Raw Debug") instead of scattered
  check‑boxes.
* **Config block** at the top: tweak expected hours, threshold, even change
  country with one line.
"""
# pylint: disable=invalid-name

from __future__ import annotations

import calendar
import itertools
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Final, Iterable

import pandas as pd
import streamlit as st
from pandas.tseries.offsets import MonthBegin

try:
    import holidays  # type: ignore
except ModuleNotFoundError as exc:  # pragma: no cover
    raise ModuleNotFoundError("pip install holidays  – required for public‑holiday look‑ups") from exc

###############################################################################
# CONFIGURATION
###############################################################################
HERE: Final = Path(__file__).parent

COUNTRY_HOLIDAYS = "EE"  # ISO‑3166 alpha‑2; swap to your country if needed
DAILY_EXPECTED_HOURS: Final[float] = 8.0  # work‑day length
LOW_PCT_THRESHOLD: Final[float] = 0.60    # red‑text threshold for % columns

# any *substring* (case‑insensitive) that indicates a paid absence
typical_keywords: tuple[str, ...] = (
    "vacation",
    "annual leave",
    "long service award",
    "military leave",
    "sick leave",
)
ABSENCE_KEYWORDS: Final[set[str]] = set(word.lower() for word in typical_keywords)

###############################################################################
# HELPER UTILITIES
###############################################################################

def _working_days(start: date, end: date, holiday_set: set[date]) -> int:
    """Return #business days (*Mon–Fri*) in [`start`, `end`] minus *holiday_set*."""
    business = pd.bdate_range(start, end, freq="C", weekmask="Mon Tue Wed Thu Fri")
    return sum(d not in holiday_set for d in business)


def working_days_month(year: int, month: int, holiday_set: set[date]) -> int:
    first = date(year, month, 1)
    last = (datetime(year, month, 1) + MonthBegin(1) - timedelta(days=1)).date()
    return _working_days(first, last, holiday_set)


def working_days_iso_week(year: int, iso_week: int, holiday_set: set[date]) -> int:
    monday = date.fromisocalendar(year, iso_week, 1)
    friday = monday + timedelta(days=4)
    return _working_days(monday, friday, holiday_set)


def is_absence(text: str | float | int | None) -> bool:  # Excel may emit floats
    txt = str(text or "").lower()
    return any(key in txt for key in ABSENCE_KEYWORDS)

###############################################################################
# DATA PIPELINE (cached)
###############################################################################

@st.cache_data(show_spinner="Parsing Excel …")
def load_xlsx(bytes_data: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(bytes_data))


@st.cache_data(show_spinner="Crunching numbers …")
def process_attendance(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return (summary_month, person_month, person_week, team_week)."""

    df = raw.copy()
    # ------------------------------------------------------------------
    # Normalise columns
    # ------------------------------------------------------------------
    df["Attendance date"] = pd.to_datetime(df["Attendance date"], format="%d.%m.%Y", errors="coerce")
    df["HoursWorked"] = pd.to_numeric(df["Total time worked decimal value"], errors="coerce").fillna(0.0)
    df["Present"] = df["HoursWorked"] > 0
    df["Vacation"] = df["Event"].map(is_absence)

    # ------------------------------------------------------------------
    # Holiday calendar
    # ------------------------------------------------------------------
    years = df["Attendance date"].dt.year.unique().tolist()
    hols = {d for y in years for d in holidays.country_holidays(COUNTRY_HOLIDAYS, years=[int(y)]).keys()}

    # Filter out weekends & public holidays (they never count toward expected)
    df = df[df["Attendance date"].dt.dayofweek < 5]
    df = df[~df["Attendance date"].dt.date.isin(hols)]

    # ------------------------------------------------------------------
    # MONTH‑LEVEL CALC
    # ------------------------------------------------------------------
    latest = df["Attendance date"].max()
    year, month = latest.year, latest.month
    workdays_month = working_days_month(year, month, hols)
    period_label = datetime(year, month, 1).strftime("%B %Y")

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
    for col_n, col_d, new in (
        ("DaysInOffice", "ExpectedDays", "PctWorkingDays"),
        ("ActualHours", "ExpectedHours", "PctHours"),
    ):
        person_month[new] = (person_month[col_n] / person_month[col_d].replace(0, pd.NA)).round(2)

    # Monthly team roll‑up
    team_size = df_month["Employee name"].nunique()
    vac_persondays_month = vac_days_month.sum()
    exp_person_days_m = workdays_month * team_size - vac_persondays_month
    summary_month = pd.DataFrame({
        "Month": [period_label],
        "Working Days": [workdays_month],
        "Team Size": [team_size],
        "Vacation Person‑Days": [vac_persondays_month],
        "Team Presence %": [(df_month["Present"].sum() / exp_person_days_m).round(2)],
        "Team Hours %": [(
            df_month["HoursWorked"].sum() / (exp_person_days_m * DAILY_EXPECTED_HOURS)
        ).round(2)],
    })

    # ------------------------------------------------------------------
    # WEEK‑LEVEL CALC (ISO calendar)
    # ------------------------------------------------------------------
    df["ISOYear"] = df["Attendance date"].dt.isocalendar().year.astype(int)
    df["ISOWeek"] = df["Attendance date"].dt.isocalendar().week.astype(int)

    base_weeks = (
        pd.DataFrame(df[["ISOYear", "ISOWeek"].drop_duplicates()])
        .assign(
            WorkingDays=lambda d: d.apply(lambda r: working_days_iso_week(r.ISOYear, r.ISOWeek, hols), axis=1)
        )
        .astype({"WorkingDays": "int8"})
        .assign(ExpectedHoursWeek=lambda d: d.WorkingDays * DAILY_EXPECTED_HOURS)
    )

    vac_days_w = (
        df[df["Vacation"]]
        .groupby(["ISOYear", "ISOWeek", "Employee name"])["Attendance date"].nunique()
        .rename("VacationDays")
    )

    person_week = (
        df.groupby(["ISOYear", "ISOWeek", "Employee name"])
        .agg(DaysInOffice=("Present", "sum"), ActualHours=("HoursWorked", "sum"))
        .join(vac_days_w, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
        .merge(base_weeks, on=["ISOYear", "ISOWeek"])
    )

    person_week["ExpectedDays"] = person_week["WorkingDays"] - person_week["VacationDays"]
    person_week["ExpectedHours"] = person_week["ExpectedDays"] * DAILY_EXPECTED_HOURS
    person_week["Year‑Week"] = person_week["ISOYear"].astype(str) + "‑W" + person_week["ISOWeek"].astype(str).str.zfill(2)

    for col_n, col_d, new in (
        ("DaysInOffice", "ExpectedDays", "PctWorkingDays"),
        ("ActualHours", "ExpectedHours", "PctHours"),
    ):
        person_week[new] = (person_week[col_n] / person_week[col_d].replace(0, pd.NA)).round(2)

    vac_pd_week = vac_days_w.groupby(["ISOYear", "ISOWeek"]).sum().rename("VacPD")

    team_week = (
        df.groupby(["ISOYear", "ISOWeek"])
        .agg(PersonDays=("Present", "sum"), ActualTeamHours=("HoursWorked", "sum"))
        .reset_index()
        .merge(base_weeks, on=["ISOYear", "ISOWeek"])
        .join(vac_pd_week, on=["ISOYear", "ISOWeek"])
        .fillna({"VacPD": 0})
        .assign(ExpectedPersonDays=lambda d: d.WorkingDays * team_size - d.VacPD)
    )

    team_week["ExpectedTeamHours"] = team_week["ExpectedPersonDays"] * DAILY_EXPECTED_HOURS
    team_week["Year‑Week"] = team_week["ISOYear"].astype(str) + "‑W" + team_week["ISOWeek"].astype(str).str.zfill(2)
    team_week["TeamPresence%"] = (
        team_week.PersonDays / team_week.ExpectedPersonDays.replace(0, pd.NA)
    ).round(2)
    team_week["TeamHours%"] = (
        team_week.ActualTeamHours / team_week.ExpectedTeamHours.replace(0, pd.NA)
    ).round(2)

    return summary_month, person_month, person_week, team_week

###############################################################################
# PRESENTATION
###############################################################################

def style_pct(df: pd.DataFrame, cols: Iterable[str]) -> pd.Styler:
    fmt = {c: "{:.0%}" for c in cols}

    def red_low(val: float):  # noqa: ANN001
        return "color:red;" if pd.notna(val) and val < LOW_PCT_THRESHOLD else ""

    return df.style.hide(axis="index").format(fmt).applymap(red_low, subset=cols)

###############################################################################
# STREAMLIT UI
###############################################################################

def main() -> None:  # pragma: no cover
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

    tab1, tab2 = st.tabs(["📈 Summary", "🪵 Raw Debug"])

    with tab1:
        st.subheader("Monthly Snapshot")
        st.dataframe(style_pct(summary_m, ["Team Presence %", "Team Hours %"]), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Per‑Person (Month)")
            st.dataframe(style_pct(person_m, ["PctWorkingDays", "PctHours"]), use_container_width=True)
        with c2:
            st.subheader("Per‑Person (Week)")
            st.dataframe(style_pct(person_w, ["PctWorkingDays", "PctHours"]), use_container_width=True)

        st.subheader("Team (Week)")
        st.dataframe(style_pct(team_w, ["TeamPresence%", "TeamHours%"]), use_container_width=True)

        st.download_button(
            "Download Monthly Summary", summary_m.to_csv(index=False).encode(), file_name="attendance_summary.csv", mime="text/csv"
        )

    with tab2:
        st.write("### Raw dataframe (after filters)")
        st.write(raw_df.head(100))


if __name__ == "__main__":  # pragma: no cover
    main()
