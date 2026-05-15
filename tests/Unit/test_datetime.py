import unittest
from datetime import datetime, date, timedelta
# Adjust the import statement below to match your file naming structure
# from my_module import DateRime, VBADate 

class TestDateRime(unittest.TestCase):

    def test_now(self) -> None:
        """Tests that Now returns the current time within a tight threshold."""
        vba_now = DateRime.Now()
        self.assertIsInstance(vba_now, VBADate)
        py_dt = vba_now.to_datetime()
        # Ensure it matches current time closely
        self.assertTrue((datetime.now() - py_dt).total_seconds() < 2)

    def test_date(self) -> None:
        """Tests that Date isolates the day component with zeroed time."""
        vba_date = DateRime.Date()
        py_dt = vba_date.to_datetime()
        now = datetime.now()
        self.assertEqual(py_dt.year, now.year)
        self.assertEqual(py_dt.month, now.month)
        self.assertEqual(py_dt.day, now.day)
        self.assertEqual(py_dt.hour, 0)
        self.assertEqual(py_dt.minute, 0)

    def test_time(self) -> None:
        """Tests that Time returns current hours/minutes with VBA base date 1899-12-30."""
        vba_time = DateRime.Time()
        py_dt = vba_time.to_datetime()
        now = datetime.now()
        self.assertEqual(py_dt.year, 1899)
        self.assertEqual(py_dt.month, 12)
        self.assertEqual(py_dt.day, 30)
        self.assertEqual(py_dt.hour, now.hour)

    def test_date_serial(self) -> None:
        """Tests manual generation of date serial values."""
        res = DateRime.DateSerial(2026, 5, 14)
        self.assertEqual(res.to_datetime(), datetime(2026, 5, 14, 0, 0))

    def test_time_serial(self) -> None:
        """Tests manual generation of time serial values with VBA base date."""
        res = DateRime.TimeSerial(14, 30, 15)
        self.assertEqual(res.to_datetime(), datetime(1899, 12, 30, 14, 30, 15))

    def test_date_value(self) -> None:
        """Tests string conversion into a valid date."""
        res = DateRime.DateValue("2026-05-14")
        self.assertEqual(res.to_datetime(), datetime(2026, 5, 14, 0, 0))

    def test_time_value(self) -> None:
        """Tests string conversion into a valid time expression."""
        res = DateRime.TimeValue("14:30:15")
        self.assertEqual(res.to_datetime(), datetime(1899, 12, 30, 14, 30, 15))

    def test_date_extraction_fields(self) -> None:
        """Tests field-specific extraction components like Day, Month, Year."""
        dt_target = VBADate(datetime(2026, 5, 14, 13, 45, 30))
        self.assertEqual(DateRime.Year(dt_target), 2026)
        self.assertEqual(DateRime.Month(dt_target), 5)
        self.assertEqual(DateRime.Day(dt_target), 14)
        self.assertEqual(DateRime.Hour(dt_target), 13)
        self.assertEqual(DateRime.Minute(dt_target), 45)
        self.assertEqual(DateRime.Second(dt_target), 30)

    def test_weekday(self) -> None:
        """Tests extraction of calendar weekday integer index values."""
        # May 14, 2026 is a Thursday
        dt_target = VBADate(datetime(2026, 5, 14))
        # Default vbSunday=1 means Sunday=1, Monday=2, Tuesday=3, Wednesday=4, Thursday=5
        self.assertEqual(DateRime.Weekday(dt_target), 5)

    def test_month_and_weekday_names(self) -> None:
        """Tests conversion of numeric date identifiers to localized string expressions."""
        self.assertEqual(DateRime.MonthName(5), "May")
        self.assertEqual(DateRime.MonthName(5, abbreviate=True), "May")
        # 5 is Thursday under default system values (Sunday=1)
        self.assertEqual(DateRime.WeekdayName(5), "Thursday")

    def test_date_add(self) -> None:
        """Tests math-based modifications across key time intervals."""
        base = VBADate(datetime(2026, 5, 14, 12, 0, 0))
        
        # Test Year increment
        self.assertEqual(DateRime.DateAdd("yyyy", 2, base).to_datetime().year, 2028)
        # Test Month increment
        self.assertEqual(DateRime.DateAdd("m", 1, base).to_datetime().month, 6)
        # Test Day increment
        self.assertEqual(DateRime.DateAdd("d", 5, base).to_datetime().day, 19)
        # Test Hour increment
        self.assertEqual(DateRime.DateAdd("h", 3, base).to_datetime().hour, 15)

    def test_date_diff(self) -> None:
        """Tests span verification calculations across arbitrary ranges."""
        d1 = VBADate(datetime(2026, 5, 14))
        d2 = VBADate(datetime(2027, 5, 14))
        d3 = VBADate(datetime(2026, 5, 24))

        self.assertEqual(DateRime.DateDiff("yyyy", d1, d2), 1)
        self.assertEqual(DateRime.DateDiff("m", d1, d2), 12)
        self.assertEqual(DateRime.DateDiff("d", d1, d3), 10)

    def test_date_part(self) -> None:
        """Tests isolated metric validation calls via string identifiers."""
        dt_target = VBADate(datetime(2026, 5, 14))
        self.assertEqual(DateRime.DatePart("yyyy", dt_target), 2026)
        self.assertEqual(DateRime.DatePart("m", dt_target), 5)
        self.assertEqual(DateRime.DatePart("d", dt_target), 14)

    def test_timer(self) -> None:
        """Tests that Timer returns numerical float duration lengths correctly."""
        sec_elapsed = DateRime.Timer()
        self.assertIsInstance(sec_elapsed, float)
        self.assertTrue(0 <= sec_elapsed <= 86400)


if __name__ == "__main__":
    unittest.main()
