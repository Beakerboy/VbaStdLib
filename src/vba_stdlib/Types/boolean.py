from __future__ import annotations
from typing import Any, Union

class VBABool:
    """
    A class to emulate VBA Boolean behavior where True == -1 and False == 0.
    """
    def __init__(self, value: Any) -> None:
        # VBA's CBool logic: 0 is False, any other numeric value is True
        self._value: bool = bool(value)

    def __bool__(self) -> bool:
        return self._value

    def __int__(self) -> int:
        # The core of VBA logic: True is -1
        return -1 if self._value else 0

    def __index__(self) -> int:
        # Allows use in bitwise ops and slices as an integer
        return self.__int__()

    def __repr__(self) -> str:
        return "True" if self._value else "False"

    def __str__(self) -> str:
        return self.__repr__()

    # Bitwise operations to match VBA's Not, And, Or
    def __invert__(self) -> VBABool:
        """Emulates VBA 'Not' (bitwise flip: Not -1 = 0)"""
        return VBABool(not self._value)

    def __and__(self, other: Union[VBABool, int, bool]) -> VBABool:
        """Emulates VBA 'And'"""
        other_val = bool(other) if not isinstance(other, int) else bool(other != 0)
        return VBABool(self._value & other_val)

    def __or__(self, other: Union[VBABool, int, bool]) -> VBABool:
        """Emulates VBA 'Or'"""
        other_val = bool(other) if not isinstance(other, int) else bool(other != 0)
        return VBABool(self._value | other_val)

    # Comparison logic
    def __eq__(self, other: Any) -> bool:
        if isinstance(other, VBABool):
            return self._value == other._value
        if isinstance(other, int):
            return int(self) == other
        return self._value == bool(other)

    # Arithmetic integration (VBA allows: 5 + True = 4)
    def __add__(self, other: int) -> int:
        return int(self) + other

    def __radd__(self, other: int) -> int:
        return self.__add__(other)
