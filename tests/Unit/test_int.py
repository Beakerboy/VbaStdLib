from vba_stdlib.literal_factory import literal_from_string


def test_boolean() -> None:
    assert literal_from_string("true")
    assert not literal_from_string("False")


def test_decimal() -> None:
    assert literal_from_string("10") == 10


def test_hex() -> None:
    assert literal_from_string("&HA") == 10


def test_oct() -> None:
    assert literal_from_string("&O12") == 10


def test_float() -> None:
    assert literal_from_string("10.") == 10.0
    assert literal_from_string("3.14") == 3.14
    assert literal_from_string(".5") == .5


def test_string() -> None:
    assert literal_from_string("foo") == "foo"
    
