from vba_stdlib.literal_factory import literal_from_string


test_decimal() -> None:
    assert literal_from_string("10") == 10


test_hex() -> None:
    assert literal_from_string("&HA") == 10


test_oct() -> None:
    assert literal_from_string("&O12") == 10
