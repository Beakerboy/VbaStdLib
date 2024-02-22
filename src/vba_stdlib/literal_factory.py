import re
from typing import Any


def literal_from_string(value: str) -> Any:
    """
    Create a VBA literal type from a string.
    """
    boolean_pattern = "TRUE|FALSE"
    hex_pattern = "&H[0-9A-F]+"
    oct_pattern = "&[Oo]?[0-7]+"
    dec_pattern = "[0-9]+"
    date_pattern = "##"
    exp = "[DEde][+-]?\d\d*"
    float_pattern1 = "\.\d\d*(" + exp + ")?"
    float_pattern2 = "\d\d*" + exp
    float_pattern3 = "\d\d\.(\d\d*)?(" + exp + ")?"
    if re.fullmatch(boolean_pattern, value.upper()):
        return value.upper() == "TRUE"
    if re.fullmatch(hex_pattern, value):
        return int("0x" + value[2:])
    if re.fullmatch(oct_pattern, value):
        start = 2 if value[1].upper() == 'O' else 1
        return int(value[start:], 8)
    if re.fullmatch(dec_pattern, value):
        return int(value)
    if (re.fullmatch(float_pattern1, value) or
            re.fullmatch(float_pattern2, value) or
            re.fullmatch(float_pattern3, value)):
        return float(value)
