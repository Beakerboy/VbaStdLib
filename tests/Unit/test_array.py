import pytest
from vba_stdlib.Types.array import VBAArray


def test_base_0_initialization():
    """Tests standard comma-separated initialization with default Base 0."""
    arr = VBAArray("apple", "banana", "cherry", base=0)
    assert arr[0] == "apple"
    assert arr[1] == "banana"
    assert arr[2] == "cherry"
    assert arr.lbound() == 0
    assert arr.ubound() == 2

def test_base_1_initialization():
    """Tests comma-separated initialization with explicit Base 1."""
    arr = VBAArray(100, 200, 300)
    assert arr[1] == 100
    assert arr[2] == 200
    assert arr[3] == 300
    assert arr.lbound() == 1
    assert arr.ubound() == 3

def test_multidimensional_custom_bounds():
    """Tests Array(1 To 2, 1 To 6) style initialization."""
    # Rows: 1 to 2, Cols: 1 to 6
    arr = VBAArray((1, 2), (1, 6))
    
    # Set and Get
    arr[1, 1] = "Top-Left"
    arr[2, 6] = "Bottom-Right"
    arr[1, 4] = "Middleish"
    
    assert arr[1, 1] == "Top-Left"
    assert arr[2, 6] == "Bottom-Right"
    assert arr[1, 4] == "Middleish"
    
    # Check bounds
    assert arr.lbound(1) == 1
    assert arr.ubound(1) == 2
    assert arr.lbound(2) == 1
    assert arr.ubound(2) == 6

def test_out_of_bounds_raises_error():
    """Ensures that accessing indices outside the defined bounds raises IndexError."""
    arr = VBAArray(1, 2, 3, base=1)
    
    with pytest.raises(IndexError, match="Subscript out of range"):
        _ = arr[0]  # Too low for base 1
        
    with pytest.raises(IndexError, match="Subscript out of range"):
        _ = arr[4]  # Too high

def test_dimension_mismatch():
    """Ensures accessing a 2D array with 1D index (or vice versa) fails."""
    arr_2d = VBAArray((1, 2), (1, 2))
    
    with pytest.raises(IndexError, match="dimension mismatch"):
        _ = arr_2d[1]  # Missing second dimension

def test_assignment_updates_value():
    """Verifies that __setitem__ actually modifies the internal data."""
    arr = VBAArray(None, None, base=0)
    arr[0] = "Modified"
    assert arr[0] == "Modified"
