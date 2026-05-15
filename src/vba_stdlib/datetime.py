from datetime import datetime, date, timedelta
from typing import Union, Optional, Any
import calendar

# Assuming VBADate is defined elsewhere in your project
# If VBADate wraps a python datetime/date object, you can adapt the conversions below.
class VBADate:
    def __init__(self, dt: datetime):
        self.dt = dt
    
    def to_datetime(self) -> datetime:
        return self.dt
    
    @classmethod
    def from_datetime(cls, dt: datetime) -> 'VBADate':
        return cls(dt)


class DateRime:
    """
    Implements MS-VBA DateTime module functions using Python's datetime library.
    Accepts and returns VBADate objects to mimic native VBA behavior.
    """

    @staticmethod
    def _to_py_dt(d: Union[VBADate, datetime, date]) -> datetime:
        """Helper to normalize inputs to a Python datetime object."""
        if isinstance(d, VBADate):
            return d.to_datetime()
        if isinstance(d, datetime):
            return d
        if isinstance(d, date):
            return datetime.combine(d, datetime.min.time())
        raise TypeError("Unsupported date type")

    @classmethod
    def Date(cls) -> VBADate:
        """Returns the current system date."""
        now = datetime.now()
        return VBADate.from_datetime(datetime(now.year, now.month, now.day))

    @classmethod
    def Time(cls) -> VBADate:
        """Returns the current system time (with a base VBA date component)."""
        now = datetime.now()
        return VBADate.from_datetime(datetime(1899, 12, 30, now.hour, now.minute, now.second))

    @classmethod
    def Now(cls) -> VBADate:
        """Returns the current system date and time."""
        return VBADate.from_datetime(datetime.now())

    @classmethod
    def DateSerial(cls, year: int, month: int, day: int) -> VBADate:
        """Returns a VBADate for a specified year, month, and day."""
        # VBA handles rolling over dates (e.g., Month 13 rolls to next year)
        # For strict matching, standard math or rolling logic would be needed.
        # This handles standard valid inputs.
        return VBADate.from_datetime(datetime(year, month, day))

    @classmethod
    def TimeSerial(cls, hour: int, minute: int, second: int) -> VBADate:
        """Returns a VBADate containing the time for a specified hour, minute, and second."""
        return VBADate.from_datetime(datetime(1899, 12, 30, hour, minute, second))

    @classmethod
    def DateValue(cls, date_string: str) -> VBADate:
        """Converts a string representation of a date to a VBADate."""
        # Simple ISO parser template; expand formats as needed for production
        dt = datetime.fromisoformat(date_string)
        return VBADate.from_datetime(datetime(dt.year, dt.month, dt.day))

    @classmethod
    def TimeValue(cls, time_string: str) -> VBADate:
        """Converts a string representation of a time to a VBADate."""
        dt = datetime.strptime(time_string, "%H:%M:%S")
        return VBADate.from_datetime(datetime(1899, 12, 30, dt.hour, dt.minute, dt.second))

    @classmethod
    def Day(cls, date_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the day of the month (1-31)."""
        return cls._to_py_dt(date_expr).day

    @classmethod
    def Month(cls, date_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the month of the year (1-12)."""
        return cls._to_py_dt(date_expr).month

    @classmethod
    def Year(cls, date_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the year."""
        return cls._to_py_dt(date_expr).year

    @classmethod
    def Hour(cls, time_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the hour (0-23)."""
        return cls._to_py_dt(time_expr).hour

    @classmethod
    def Minute(cls, time_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the minute (0-59)."""
        return cls._to_py_dt(time_expr).minute

    @classmethod
    def Second(cls, time_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the second (0-59)."""
        return cls._to_py_dt(time_expr).second

    @classmethod
    def Weekday(cls, date_expr: Union[VBADate, datetime, date], first_day_of_week: int = 1) -> int:
        """
        Returns the weekday integer.
        VBA default (vbSunday = 1, vbMonday = 2, ..., vbSaturday = 7).
        """
        py_dow = cls._to_py_dt(date_expr).weekday()  # Python: Mon=0, Sun=6
        vba_dow = (py_dow + 1) % 7 + 1               # Shift to Sun=1, Mon=2
        
        # Adjust based on first_day_of_week if custom logic is requested
        shift = first_day_of_week - 1
        adjusted = (vba_dow - shift - 1) % 7 + 1
        return adjusted

    @classmethod
    def MonthName(cls, month: int, abbreviate: bool = False) -> str:
        """Returns the string name of the specified month."""
        if abbreviate:
            return calendar.month_abbr[month]
        return calendar.month_name[month]

    @classmethod
    def WeekdayName(cls, weekday: int, abbreviate: bool = False, first_day_of_week: int = 1) -> str:
        """Returns the string name of the specified weekday."""
        # Convert VBA weekday index to Python calendar index (Mon=0...Sun=6)
        # Assuming default vbSunday=1
        vba_to_py = [6, 0, 1, 2, 3, 4, 5]
        py_day = vba_to_py[(weekday - 1 + (first_day_of_week - 1)) % 7]
        
        if abbreviate:
            return calendar.day_abbr[py_day]
        return calendar.day_name[py_day]

    @classmethod
    def DateAdd(cls, interval: str, number: int, date_expr: Union[VBADate, datetime, date]) -> VBADate:
        """Adds a specific time interval to a date."""
        dt = cls._to_py_dt(date_expr)
        interval = interval.lower()

        if interval == "yyyy":
            res = dt.replace(year=dt.year + number)
        elif interval == "q":
            res = dt + timedelta(days=number * 90) # Approximation
        elif interval == "m":
            month = dt.month - 1 + number
            year = dt.year + month // 12
            month = month % 12 + 1
            day = min(dt.day, calendar.monthrange(year, month)[1])
            res = dt.replace(year=year, month=month, day=day)
        elif interval in ("y", "d", "w"):
            res = dt + timedelta(days=number)
        elif interval == "ww":
            res = dt + timedelta(weeks=number)
        elif interval == "h":
            res = dt + timedelta(hours=number)
        elif interval == "n":
            res = dt + timedelta(minutes=number)
        elif interval == "s":
            res = dt + timedelta(seconds=number)
        else:
            raise ValueError(f"Invalid interval: {interval}")

        return VBADate.from_datetime(res)

    @classmethod
    def DateDiff(cls, interval: str, date1: Union[VBADate, datetime, date], date2: Union[VBADate, datetime, date]) -> int:
        """Returns the number of time intervals between two dates."""
        dt1 = cls._to_py_dt(date1)
        dt2 = cls._to_py_dt(date2)
        diff = dt2 - dt1
        interval = interval.lower()

        if interval == "yyyy":
            return dt2.year - dt1.year
        elif interval == "q":
            return (dt2.year - dt1.year) * 4 + (dt2.month - dt1.month) // 3
        elif interval == "m":
            return (dt2.year - dt1.year) * 12 + (dt2.month - dt1.month)
        elif interval in ("y", "d"):
            return diff.days
        elif interval == "w":
            return diff.days // 7
        elif interval == "ww":
            # Calendar weeks between dates
            return diff.days // 7 
        elif interval == "h":
            return int(diff.total_seconds() // 3600)
        elif interval == "n":
            return int(diff.total_seconds() // 60)
        elif interval == "s":
            return int(diff.total_seconds())
        else:
            raise ValueError(f"Invalid interval: {interval}")

    @classmethod
    def DatePart(cls, interval: str, date_expr: Union[VBADate, datetime, date]) -> int:
        """Returns the specified part of a given date."""
        dt = cls._to_py_dt(date_expr)
        interval = interval.lower()

        if interval == "yyyy":
            return dt.year
        elif interval == "q":
            return (dt.month - 1) // 3 + 1
        elif interval == "m":
            return dt.month
        elif interval == "y":
            return dt.timetuple().tm_yday
        elif interval == "d":
            return dt.day
        elif interval == "w":
            return cls.Weekday(dt)
        elif interval == "ww":
            return dt.isocalendar()[1]
        elif interval == "h":
            return dt.hour
        elif interval == "n":
            return dt.minute
        elif interval == "s":
            return dt.second
        else:
            raise ValueError(f"Invalid interval: {interval}")

    @classmethod
    def Timer(cls) -> float:
        """Returns the number of seconds elapsed since midnight."""
        now = datetime.now()
        midnight = datetime.combine(now.date(), datetime.min.time())
        return (now - midnight).total_seconds()
