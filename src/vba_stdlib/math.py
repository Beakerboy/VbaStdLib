from __future__ import annotations
from vba_types import VBADouble, VBATypeBase


class Math:

    @staticmethod
    def abs(number: VBATypeBase) -> VBADouble:
        # if isinstance(number, Null):
        #     return number
        # if isinstance(number, Empty):
        #     return 0
        return VBADouble(abs(number.value))
