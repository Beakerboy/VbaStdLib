from vba_stdlib.Types.empty import Empty
from vba_stdlib.Types.null import Null
from typing import Any, TypeVar


class Math:

    @staticmethod
    def abs(number: Any) -> Any:
        if isinstance(number, Null):
            return number
        if isinstance(number, Empty):
            return 0
        return abs(number)

    @staticmethod
    def atn(number: Any) -> number:
        pass

    @staticmethod
    def cos(number: Any) -> number:
        pass
    
    @staticmethod
    def exp(number: Any) -> number:
        pass

    @staticmethod
    def log(number: Any) -> number:
        pass

    @staticmethod
    def rnd(number: Any) -> number:
        pass

    @staticmethod
    def round(number: Any) -> number:
        pass

    @staticmethod
    def sgn(number: Any) -> number:
        pass

    @staticmethod
    def sin(number: Any) -> number:
        pass

    @staticmethod
    def Sqr(number: Any) -> number:
        pass

    @staticmethod
    def tan(number: Any) -> number:
        pass

    @staticmethod
    def randomize(number: Any) -> None:
        pass
