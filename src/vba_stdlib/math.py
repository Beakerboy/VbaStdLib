from vba_stdlib.Types.empty import Empty
from vba_stdlib.Types.null import Null


class Math:

    def abs(number: Any) -> Any:
        if isinstance(number, Null):
            return number
        if isinstance(number, Empty):
            return 0
        return abs(number)

    def atn(number: Any) -> number:
        pass

    def cos(number: Any) -> number:
        pass

    def exp(number: Any) -> number:
        pass

    def log(number: Any) -> number:
        pass

    def rnd(number: Any) -> number:
        pass

    def round(number: Any) -> number:
        pass

    def sgn(number: Any) -> number:
        pass

    def sin(number: Any) -> number:
        pass

    def Sqr(number: Any) -> number:
        pass

    def tan(number: Any) -> number:
        pass

    def randomize(number: Any) -> None:
        pass
