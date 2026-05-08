from typing import Any, Tuple, Union, List

class VBAArray:
    def __init__(self, *args: Any, base: int = 0):
        """
        Initializes a VBA-style array.
        - VBAArray(1, 2, 3) -> Base 0/1 list of values.
        - VBAArray((1, 2), (1, 6)) -> 2D array with specific LBound and UBound.
        """
        # Case 1: Tuple definitions for dimensions (e.g., (1, 2), (1, 6))
        if args and all(isinstance(arg, tuple) and len(arg) == 2 for arg in args):
            self._bounds = list(args)
            shape = tuple(max_idx - min_idx + 1 for min_idx, max_idx in self._bounds)
            self._data = self._recursive_init(shape)
        # Case 2: Comma separated list of values
        else:
            self._data = list(args)
            self._bounds = [(base, base + len(args) - 1)]
            
    def _recursive_init(self, shape: Tuple[int, ...]) -> Any:
        if len(shape) == 1:
            return [None] * shape[0]
        return [self._recursive_init(shape[1:]) for _ in range(shape[0])]

    def _get_coords(self, indices: Tuple[int, ...]) -> Tuple[int, ...]:
        if len(indices) != len(self._bounds):
            raise IndexError("Subscript out of range (dimension mismatch)")
        
        internal = []
        for i, idx in enumerate(indices):
            low, high = self._bounds[i]
            if not (low <= idx <= high):
                raise IndexError(f"Subscript out of range: {idx} (Expected {low} to {high})")
            internal.append(idx - low)
        return tuple(internal)

    def __getitem__(self, key: Union[int, Tuple[int, ...]]) -> Any:
        indices = key if isinstance(key, tuple) else (key,)
        coords = self._get_coords(indices)
        val = self._data
        for c in coords:
            val = val[c]
        return val

    def __setitem__(self, key: Union[int, Tuple[int, ...]], value: Any) -> None:
        indices = key if isinstance(key, tuple) else (key,)
        coords = self._get_coords(indices)
        target = self._data
        for c in coords[:-1]:
            target = target[c]
        target[coords[-1]] = value

    def lbound(self, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][0]

    def ubound(self, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][1]

    def __repr__(self) -> str:
        return f"<VBAArray: Bounds {self._bounds}>"
style!
print(arr2.ubound(2))      # 6
