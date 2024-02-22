def literal_from_string(value: str) -> Any:
    """
    Create a VBA literal type from a string.
    """
    boolean_pattern = "TRUE|FALSE"
    hex_pattern = "&H[0-9A-F]+"
    oct_pattern = "&[Oo]?[0-7]+"
    dec_pattern = "[0-9]+"
    if re.search(boolean_pattern, value.upper()):
        return value.upper() == "TRUE"
    if re.search(hex_pattern, value):
        return int("0x" + value[2:])
    if re.search(oct_pattern, value):
        start = 2 if value[1].upper() == 'O' else 1
        return int(value[start:], 8)
    if re.search(dec_pattern, value):
        return int(value)
