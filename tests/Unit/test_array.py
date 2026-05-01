from vba_stdlib.Types.array import Array


def test_initialization_with_base():
    """Test if Array.base correctly mimics Option Base 1/0."""
    Array.base = 1
    arr1 = Array(5) # Expect 1 To 5
    assert arr1._bounds[0] == (1, 5)
    
    Array.base = 0
    arr0 = Array(5) # Expect 0 To 5
    assert arr0._bounds[0] == (0, 5)

def test_explicit_tuple_bounds():
    """Test Dim A(1 To 3, 6 To 9) equivalent."""
    arr = Array((1, 3), (6, 9))
    arr[1, 6] = "top-left"
    arr[3, 9] = "bottom-right"
    assert arr[1, 6] == "top-left"
    assert arr[3, 9] == "bottom-right"

def test_subscript_out_of_range():
    """Ensure accessing outside bounds raises IndexError."""
    arr = Array((1, 5))
    with pytest.raises(IndexError):
        _ = arr[0]
    with pytest.raises(IndexError):
        _ = arr[6]

def test_redim_clears_data():
    """Standard ReDim should reset all elements to None."""
    arr = Array(1, 2)
    arr[1] = "data"
    arr.redim(1, 5)
    assert arr[1] is None

def test_redim_preserve_1d():
    """Test ReDim Preserve for 1D arrays."""
    arr = Array(1, 2)
    arr[1] = "keep"
    arr.redim(1, 10, preserve=True)
    assert arr[1] == "keep"
    assert arr[10] is None

def test_redim_preserve_multidim_restriction():
    """VBA only allows changing the LAST dimension during Preserve."""
    arr = Array((1, 2), (1, 2))
    with pytest.raises(ValueError, match="Preserve only allows changing the last dimension"):
        # Attempting to change the first dimension (1 To 2 -> 1 To 3)
        arr.redim((1, 3), (1, 2), preserve=True)

def test_redim_preserve_last_dim_success():
    """Test successful ReDim Preserve on the second dimension."""
    arr = Array((1, 2), (1, 2))
    arr[2, 2] = "target"
    arr.redim((1, 2), (1, 10), preserve=True)
    assert arr[2, 2] == "target"
    assert arr[2, 10] is None

def test_invalid_bound_values():
    """Ensure LBound > UBound raises ValueError."""
    with pytest.raises(ValueError):
        Array((10, 1))
