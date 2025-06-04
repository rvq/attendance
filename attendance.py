import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO
import holidays  # Estonia national holidays

###############################################################################
# CONFIG
###############################################################################

DAILY_EXPECTED_HOURS = 8.0  # tweak if a full work-day ≠ 8 h
EE_HOLIDAYS = holidays.EE()  # public holidays for every year

###############################################################################
# Helper functions
###############################################################################

def working_days_in_month(year: int, month: int) -> int:
    """Mon–Fri days in *this* month, excluding public holidays."""
    cal = calendar.Calendar()
    return sum(
        1
        for day, dow in cal.itermonthdays2(year, month)
        if day and dow < 5 and datetime(year, month, day).date() not in EE_HOLIDAYS
    )

###############################################################################
# Core processor
###############################################################################

def process_attendance(df: pd.DataFrame):
    """Return all attendance & hour metrics.

    Required columns (exact case):
      • Employee name
      • Attendance date               (dd.mm.yyyy)
      • Total time worked decimal value (float hours, blanks ok)
      • Event                          ("Vacation" marks annual leave)
    """

    df = df.copy()
    df["Attendance date"] = pd.to_datetime(df["Attendance date"], format="%d.%m.%Y")

    # Infer hours worked (0 for blanks) + simple presence flag
    df["HoursWorked"] = pd.to_numeric(
        df["Total time worked decimal value"], errors="coerce"
    ).fillna(0.0)
    df["Present"] = df["HoursWorked"] > 0

    # Detect vacations (case-insensitive ‘Vacation’ in Event column)
    df["Vacation"] = (
        df["Event"].fillna("").str.strip().str.lower() == "vacation"
    )

    # Filter out weekends and public holidays. Vacations are kept (they need counting)
    df = df[df["Attendance date"].dt.dayofweek < 5]
    df = df[~df["Attendance date"].dt.date.isin(EE_HOLIDAYS)]

    # Latest month present in the sheet drives the report window
    latest_date = df["Attendance date"].max()
    year, month = latest_date.year, latest_date.month
    ym_period = pd.Period(datetime(year, month, 1), "M")

    # ------------------------------------------------------------------
    # MONTH-LEVEL CONSTANTS
    # ------------------------------------------------------------------
    working_days_month = working_days_in_month(year, month)

    # Slice month rows once
    month_mask = (
        (df["Attendance date"].dt.year == year)
        & (df["Attendance date"].dt.month == month)
    )
    df_month = df[month_mask]

    # ------------------------------------------------------------------
    # PER-PERSON · MONTH
    # ------------------------------------------------------------------
    # Vacation days per person this month
    vac_days_month = (
        df_month.groupby("Employee name")["Vacation"].sum().rename("VacationDays")
    )

    person_month = (
        df_month.groupby("Employee name")
        .agg(
            DaysInOffice=("Present", "sum"),
            ActualHours=("HoursWorked", "sum"),
        )
        .join(vac_days_month, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
    )

    person_month["ExpectedDays"] = (
        working_days_month - person_month["VacationDays"]
    )
    person_month["ExpectedHours"] = (
        person_month["ExpectedDays"] * DAILY_EXPECTED_HOURS
    )

    # Guard against division by zero
    person_month["PctOfWorkingDays"] = (
        person_month["DaysInOffice"] / person_month["ExpectedDays"].replace(0, pd.NA)
    ).round(2)
    person_month["PctOfHours"] = (
        person_month["ActualHours"] / person_month["ExpectedHours"].replace(0, pd.NA)
    ).round(2)

    # ------------------------------------------------------------------
    # PER-PERSON · WEEK
    # ------------------------------------------------------------------
    df["ISOWeek"] = df["Attendance date"].dt.isocalendar().week.astype(int)

    # Working & vacation days by week (per person)
    week_working_days = (
        df.groupby("ISOWeek")["Attendance date"].nunique()
        .rename("WorkingDays")
        .reset_index()
    )
    week_working_days["ExpectedHours"] = (
        week_working_days["WorkingDays"] * DAILY_EXPECTED_HOURS
    )

    vac_days_week = (
        df.groupby(["ISOWeek", "Employee name"])["Vacation"].sum().rename("VacationDays")
    )

    person_week = (
        df.groupby(["ISOWeek", "Employee name"])
        .agg(
            DaysInOffice=("Present", "sum"),
            ActualHours=("HoursWorked", "sum"),
        )
        .join(vac_days_week, how="left")
        .fillna({"VacationDays": 0})
        .reset_index()
        .merge(week_working_days, on="ISOWeek")
    )

    person_week["ExpectedDays"] = (
        person_week["WorkingDays"] - person_week["VacationDays"]
    )
    person_week["ExpectedHours"] = person_week["ExpectedDays"] * DAILY_EXPECTED_HOURS
    person_week["PctOfWorkingDays"] = (
        person_week["DaysInOffice"] / person_week["ExpectedDays"].replace(0, pd.NA)
    ).round(2)
    person_week["PctOfHours"] = (
        person_week["ActualHours"] / person_week["ExpectedHours"].replace(0, pd.NA)
    ).round(2)

    # ------------------------------------------------------------------
    # TEAM-LEVEL METRICS
    # ------------------------------------------------------------------
    team_size_total = df["Employee name"].nunique()

    # ---- Weekly team presence/hours ----
    vac_persondays_week = (
        df.groupby("ISOWeek")["Vacation"].sum().rename("VacationPersonDays")
    )

    team_week = (
        df.groupby("ISOWeek")
        .agg(
            Pe