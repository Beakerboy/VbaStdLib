class VBAError(Exception):
    """Base class for VBA-style errors."""
    def __init__(self, number, message):
        self.number = number
        self.message = f"Run-time error '{number}': {message}"
        super().__init__(self.message)

class SubscriptOutOfRange(VBAError):
    """Error 9: Occurs when an index is outside array bounds."""
    def __init__(self):
        super().__init__(9, "Subscript out of range")

class ArrayLockedError(VBAError):
    """Error 10: Occurs when attempting to ReDim a locked array."""
    def __init__(self):
        super().__init__(10, "This array is fixed or temporarily locked")
