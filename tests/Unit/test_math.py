from vba_stdlib.math import math
from vba_types import VBADouble, VBAInteger


def test_abs() -> None:
    input = VBAInteger(1)
    assert math.abs(input) == input
    input2 = VBAInteger(-1)
    assert math.abs(input2) == input
