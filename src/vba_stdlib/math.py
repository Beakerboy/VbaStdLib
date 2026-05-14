from __future__ import annotations
from typing import Any, TypeVar
from vba_types import VBADouble, VBATypeBase


class Math:

    @staticmethod
    def abs(number: VBATypeBase) -> VBADouble:
        # if isinstance(number, Null):
        #     return number
        # if isinstance(number, Empty):
        #     return 0
        return VBADouble(abs(number.value))

    @staticmethod
    def atn(number: Any) -> float:
        pass

    @staticmethod
    def cos(number: Any) -> float:
        pass
    
    @staticmethod
    def exp(number: Any) -> float:
        pass

    @staticmethod
    def log(number: Any) -> float:
        pass

    @staticmethod
    def rnd(number: Any) -> float:
        pass

    @staticmethod
    def round(number: Any) -> float:
        pass

    @staticmethod
    def sgn(number: Any) -> float:
        pass

    @staticmethod
    def sin(number: Any) -> float:
        pass

    @staticmethod
    def Sqr(number: Any) -> float:
        pass

    @staticmethod
    def tan(number: Any) -> float:
        pass

    @staticmethod
    def randomize(number: Any) -> None:
        pass
