from vba_stdlib.Types.empty import Empty
from vba_stdlib.Types.null import Null
from typing import Any, TypeVar


class Math:

    @staticmethod
    def abs(number: Any) -> float:
        if isinstance(number, Null):
            return number
        if isinstance(number, Empty):
            return 0
        return abs(number)

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
