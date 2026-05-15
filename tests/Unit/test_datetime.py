import pytest
from datetime import datetime, date, timedelta
# Adjust this import to match your file and project structure
# from my_module import DateRime, VBADate


# ==========================================
# FIXTURES (Shared Test Setup)
# ==========================================

@pytest.fixture
def target_date() -> VBADate:
    """Provides a fixed, standard date for extraction and calculation tests."""
    return VBADate(datetime(2026, 5, 14, 13, 45, 30))


@pytest.fixture
def base_date() -> VBADate:
    """Provides a flat baseline date with no time components."""
    return VBADate(datetime(2026, 5, 14, 0, 0, 0))


# ==========================================
# TEST CASES
# ==========================================

def test_now() -> None:
    """Verifies that Now returns the current time within a narrow execution window."""
    vba_now = DateRime.Now()
    assert isinstance(vba_now, VBADate)
    
    py_dt = vba_now.to_datetime()
    time_delta = datetime.now() - py_dt
    assert time_delta.total_seconds() < 2


def test_date() -> None:
    """Ensures Date isolates the day component with zeroed out time parameters."""
    vba_date = DateRime.Date()
    py_dt = vba_date.to_datetime()
    now = datetime.now()
    
    assert py_dt.year == now.year
    assert py_dt.month == now.month
    assert py_dt.day == now.day
    assert py_dt.hour == 0
    assert py_dt.minute == 0


def test_time() -> None:
    """Validates that Time uses the standard VBA base date of 1899-12-30."""
    vba_time = DateRime.Time()
    py_dt = vba_time.to_datetime()
    now = datetime.now()
    
    assert py_dt.year == 1899
    assert py_dt.month == 12
    assert py_dt.day == 30
    assert py_dt.hour == now.hour


def test_date_serial() -> None:
    """Checks the assembly of manual year, month, and day components."""
    res = DateRime.DateSerial(2026, 5, 14)
    assert res.to_datetime() == datetime(2026, 5, 14, 0, 0)


def test_time_serial() -> None:
    """Checks the assembly of hours, minutes, and seconds into a base VBA date."""
    res = DateRime.TimeSerial(14, 30, 15)
    assert res.to_datetime() == datetime(1899, 12, 30, 14, 30, 15)


def test_date_value() -> None:
    """Confirms string parsing translates cleanly into a date wrapper."""
    res = DateRime.DateValue("2026-05-14")
    assert res.to_datetime() == datetime(2026, 5, 14, 0, 0)


def test_time_value() -> None:
    """Confirms time-string parsing outputs the correct time metrics."""
    res = DateRime.TimeValue("14:30:15")
    assert res.to_datetime() == datetime(1899, 12, 30, 14, 30, 15)


def test_date_extraction_fields(target_date: VBADate) -> None:
    """Verifies granular fields are extracted correctly from a VBADate object."""
    assert DateRime.Year(target_date) == 2026
    assert DateRime.Month(target_date) == 5
    assert DateRime.Day(target_date) == 14
    assert DateRime.Hour(target_date) == 13
    assert DateRime.Minute(target_date) == 45
    assert DateRime.Second(target_date) == 30


def test_weekday(base_date: VBADate) -> None:
    """Tests weekday translation (May 14, 2026 is a Thursday -> 5 under default Sun=1)."""
    assert DateRime.Weekday(base_date) == 5


def test_month_and_weekday_names() -> None:
    """Tests string literal outputs for localized month and weekday lookups."""
    assert DateRime.MonthName(5) == "May"
    assert DateRime.MonthName(5, abbreviate=True) == "May"
    assert DateRime.WeekdayName(5) == "Thursday"


def test_date_add(base_date: VBADate) -> None:
    """Ensures DateAdd mathematical offsets return expected target dates."""
    assert DateRime.DateAdd("yyyy", 2, base_date).to_datetime().year == 2028
    assert DateRime.DateAdd("m", 1, base_date).to_datetime().month == 6
    assert DateRime.DateAdd("d", 5, base_date).to_datetime().day == 19
    assert DateRime.DateAdd("h", 3, base_date).to_datetime().hour == 3


def test_date_diff(base_date: VBADate) -> None:
    """Tests the span evaluation across variable interval targets."""
    future_year = VBADate(datetime(2027, 5, 14))
    future_days = VBADate(datetime(2026, 5, 24))

    assert DateRime.DateDiff("yyyy", base_date, future_year) == 1
    assert DateRime.DateDiff("m", base_date, future_year) == 12
    assert DateRime.DateDiff("d", base_date, future_days) == 10


def test_date_part(base_date: VBADate) -> None:
    """Tests parsing specific string segments out of a target date."""
    assert DateRime.DatePart("yyyy", base_date) == 2026
    assert DateRime.DatePart("m", base_date) == 5
    assert DateRime.DatePart("d", base_date) == 14


def test_timer() -> None:
    """Validates the daily countdown tick value yields an expected floating point scalar."""
    seconds_elapsed = DateRime.Timer()
    assert isinstance(seconds_elapsed, float)
    assert 0.0 <= seconds_elapsed <= 86400.0
