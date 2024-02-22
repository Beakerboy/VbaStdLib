import re
from dateutil.parser import parse
from typing import Any


def literal_from_string(value: str) -> Any:
    """
    Create a VBA literal type from a string.
    """
    boolean_pattern = "TRUE|FALSE"
    hex_pattern = "&H[\dA-F]+"
    oct_pattern = "&[Oo]?[0-7]+"
    dec_pattern = "\d+"
    date_pattern = "##"
    exp = "[DEde][+-]?\d\d*"
    float_pattern1 = "\.\d\d*(" + exp + ")?"
    float_pattern2 = "\d\d*" + exp
    float_pattern3 = "\d\d*\.(\d\d*)?(" + exp + ")?"

    # Date Literal
    sep = "([ ]+|([]*[/-][]*))"
    eng_month = "(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)"
    month_abv = "(JAN|FEB|MAR|APR|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    month_name = "(eng_month|month_abv)"
    date_val_part = "(\d+|" + month_name + ")"
    date = date_val_part + sep + date_val_part + "(" + sep + date_val_part + ")?"
    time_pattern1 = "(\d+(AM|PM|A|P))"
    time_pattern2 = "(\d+[ ]?[:.][ ]?\d+([ ]?[:.][ ]?\d+)?)"
    time_pattern = "(" + time_pattern1 + "|" + time_pattern2 + ")"
    date_pattern = date + "([ ]+" + time_pattern + ")?"
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
    if value[-1] == '"' and value[0] == '"':
        return value[1:-1]
    if re.fullmatch("#" + date_pattern "#", value) or
            re.fullmatch("#" + time_pattern + "#", value):
        return parse(value[1:-1])
    # assume non-quoted string.
    return value
