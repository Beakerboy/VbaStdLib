from vba_stdlib.math import Math
from vba_types import VBADouble, VBAInteger


def test_abs() -> None:
    input = VBAInteger(1)
    result = Math.abs(input)
    assert result == input
    assert isinstance(result, VBADouble)
    input2 = VBAInteger(-1)
    assert Math.abs(input2) == input
