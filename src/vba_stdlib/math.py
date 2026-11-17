from __future__ import annotations
import math
import random
from typing import Any, Optional
from vba_types import VBADouble, VBATypeBase


class Math:

    # The internal seed state tracker for tracking VBA Rnd states
    _last_rnd_value: float = 0.7055475  # VBA's default initial unseeded baseline

    @staticmethod
    def abs(number: VBATypeBase) -> VBADouble:
        """Returns the absolute value of a number."""
        return VBADouble(abs(number.value))

    @staticmethod
    def atn(number: VBATypeBase) -> VBADouble:
        """Returns the arctangent of a number as a value in radians."""
        return VBADouble(math.atan(number.value))

    @staticmethod
    def cos(number: VBATypeBase) -> VBADouble:
        """Returns the cosine of an angle specified in radians."""
        return VBADouble(math.cos(number.value))

    @staticmethod
    def exp(number: VBATypeBase) -> VBADouble:
        """Returns e (the base of natural logarithms) raised to a power."""
        return VBADouble(math.exp(number.value))

    @staticmethod
    def int(number: VBATypeBase) -> VBADouble:
        """
        Returns the integer portion of a number.
        If negative, returns the first negative integer less than or equal to number.
        """
        return VBADouble(float(math.floor(number.value)))

    @staticmethod
    def fix(number: VBATypeBase) -> VBADouble:
        """
        Returns the integer portion of a number.
        If negative, returns the first negative integer greater than or equal to number.
        """
        return VBADouble(float(math.trunc(number.value)))

    @staticmethod
    def log(number: VBATypeBase) -> VBADouble:
        """Returns the natural logarithm of a number."""
        if number.value <= 0:
            raise ValueError("VBA Log function argument must be greater than zero.")
        return VBADouble(math.log(number.value))

    @classmethod
    def rnd(cls, number: Optional[VBATypeBase] = None) -> VBADouble:
        """
        Returns a pseudo-random single-precision floating point number less than 1 but >= 0.
        
        VBA Rnd behavior matrix:
        - Less than 0: Returns the same number every time, using 'number' as the seed.
        - Greater than 0: Returns the next random number in the sequence.
        - Equal to 0: Returns the most recently generated random number.
        - Not supplied: Returns the next random number in the sequence.
        """
        if number is None:
            cls._last_rnd_value = random.random()
        else:
            val = number.value
            if val < 0:
                random.seed(val)
                cls._last_rnd_value = random.random()
            elif val > 0:
                cls._last_rnd_value = random.random()
            elif val == 0:
                # Returns the exact last value generated without progressing sequence
                pass

        return VBADouble(cls._last_rnd_value)

    @staticmethod
    def sgn(number: VBATypeBase) -> VBADouble:
        """
        Returns an integer variant indicating the sign of a number.
        - Returns 1 if number > 0
        - Returns 0 if number == 0
        - Returns -1 if number < 0
        """
        val = number.value
        if val > 0:
            return VBADouble(1.0)
        elif val < 0:
            return VBADouble(-1.0)
        return VBADouble(0.0)

    @staticmethod
    def sin(number: VBATypeBase) -> VBADouble:
        """Returns the sine of an angle specified in radians."""
        return VBADouble(math.sin(number.value))

    @staticmethod
    def sqr(number: VBATypeBase) -> VBADouble:
        """Returns the square root of a number."""
        if number.value < 0:
            raise ValueError("VBA Sqr function argument cannot be negative.")
        return VBADouble(math.sqrt(number.value))

    @staticmethod
    def tan(number: VBATypeBase) -> VBADouble:
        """Returns the tangent of an angle specified in radians."""
        return VBADouble(math.tan(number.value))
