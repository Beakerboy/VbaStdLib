from typing import List, Tuple, Any, Union, Optional

class VBAArray:
    def __init__(self, bounds: List[Tuple[int, int]], data: List[Any]):
        self._bounds: List[Tuple[int, int]] = bounds
        self._data: List[Any] = data

    @classmethod
    def from_list(cls, source_list: Union[List[Any], Tuple[Any, ...]]) -> 'VBAArray':
        """Emulates VBA: arr = Array(1, 2, 3)"""
        bounds = [(0, len(source_list) - 1)]
        return cls(bounds, list(source_list))

    @classmethod
    def with_bounds(cls, *bounds_args: int) -> 'VBAArray':
        """Emulates VBA: Dim arr(1 To 4, 5 To 9)"""
        if len(bounds_args) % 2 != 0:
            raise ValueError("Bounds must be provided as pairs (lower, upper).")
        
        bounds: List[Tuple[int, int]] = []
        shape: List[int] = []
        for i in range(0, len(bounds_args), 2):
            lower, upper = bounds_args[i], bounds_args[i+1]
            bounds.append((lower, upper))
            shape.append(upper - lower + 1)
            
        def create_data(dims: List[int]) -> List[Any]:
            if len(dims) == 1:
                return [None] * dims[0]
            return [create_data(dims[1:]) for _ in range(dims[0])]
            
        return cls(bounds, create_data(shape))

    def _get_internal_indices(self, keys: Union[int, Tuple[int, ...]]) -> List[int]:
        key_tuple = keys if isinstance(keys, tuple) else (keys,)
        
        if len(key_tuple) != len(self._bounds):
            raise IndexError("Subscript out of range: Dimension mismatch.")
        
        indices = []
        for i, key in enumerate(key_tuple):
            lower, upper = self._bounds[i]
            if not (lower <= key <= upper):
                raise IndexError(f"Subscript out of range: {key}")
            indices.append(key - lower)
        return indices

    def __getitem__(self, keys: Union[int, Tuple[int, ...]]) -> Any:
        indices = self._get_internal_indices(keys)
        item = self._data
        for idx in indices:
            item = item[idx]
        return item

    def __setitem__(self, keys: Union[int, Tuple[int, ...]], value: Any) -> None:
        indices = self._get_internal_indices(keys)
        target = self._data
        for idx in indices[:-1]:
            target = target[idx]
        target[indices[-1]] = value

    def LBound(self, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][0]

    def UBound(self, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][1]

    def ReDimPreserve(self, *new_bounds: int) -> None:
        """Only the last dimension's upper bound can change."""
        new_b = [(new_bounds[i], new_bounds[i+1]) for i in range(0, len(new_bounds), 2)]
        
        if len(new_b) != len(self._bounds):
            raise ValueError("Cannot change dimensions with ReDim Preserve.")
        
        for i in range(len(new_b) - 1):
            if new_b[i] != self._bounds[i]:
                raise ValueError("Only the last dimension can be modified.")
        
        if new_b[-1][0] != self._bounds[-1][0]:
            raise ValueError("Cannot change the lower bound with ReDim Preserve.")

        new_upper = new_b[-1][1]
        new_size = new_upper - new_b[-1][0] + 1
        
        def resize_recursive(data: List[Any], depth: int) -> None:
            if depth < len(self._bounds) - 1:
                for sublist in data:
                    resize_recursive(sublist, depth + 1)
            else:
                current_size = len(data)
                if new_size > current_size:
                    data.extend([None] * (new_size - current_size))
                else:
                    del data[new_size:]

        resize_recursive(self._data, 0)
        self._bounds[-1] = (self._bounds[-1][0], new_upper)

    def __repr__(self) -> str:
        b_str = ", ".join([f"{b[0]} To {b[1]}" for b in self._bounds])
        return f"VBAArray({b_str})"
