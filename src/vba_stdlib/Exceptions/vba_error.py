from typing import TypeVar


T = TypeVar('T', bound='VBAError')


class VBAError(Exception):
    """Base class for VBA-style errors."""
    def __init__(self: T, number: int, message: str) -> None:
        self.number = number
        self.message = f"Run-time error '{number}': {message}"
        super().__init__(self.message)


T = TypeVar('T', bound='SubscriptOutOfRange')


class SubscriptOutOfRange(VBAError):
    """Error 9: Occurs when an index is outside array bounds."""
    def __init__(self: T) -> None:
        super().__init__(9, "Subscript out of range")


T = TypeVar('T', bound='ArrayLockedError')


class ArrayLockedError(VBAError):
    """Error 10: Occurs when attempting to ReDim a locked array."""
    def __init__(self: T) -> None:
        super().__init__(10, "This array is fixed or temporarily locked")
