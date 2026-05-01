import itertools
from typing import TypeVar


T = TypeVar('T', bound='Array')


class Array:
    base = 0  # Global setting mimicking 'Option Base'

    def __init__(self: T, *bounds) -> None:
        """
        Initializes an N-dimensional array.
        Examples:
            Array(10) -> Dim A(base To 10)
            Array((1, 3), (6, 9)) -> Dim A(1 To 3, 6 To 9)
        """
        self._bounds = []
        shape = []
        
        for b in bounds:
            l, u = b if isinstance(b, tuple) else (Array.base, b)
            if l > u:
                raise ValueError(f"LBound ({l}) cannot be greater than UBound ({u})")
            self._bounds.append((l, u))
            shape.append(u - l + 1)
        
        self._shape = tuple(shape)
        self._data = self._build_nested_lists(self._shape)

    def _build_nested_lists(self: T, shape):
        if len(shape) == 1:
            return [None] * shape[0]
        return [self._build_nested_lists(shape[1:]) for _ in range(shape[0])]

    def _resolve_indices(self: T, keys):
        indices = keys if isinstance(keys, tuple) else (keys,)
        if len(indices) != len(self._bounds):
            raise IndexError("Wrong number of dimensions")
            
        internal_indices = []
        for i, val in enumerate(indices):
            l, u = self._bounds[i]
            if not (l <= val <= u):
                raise IndexError(f"Subscript out of range: Dim {i+1}")
            internal_indices.append(val - l)
        return internal_indices

    def __getitem__(self: T, keys):
        target = self._data
        for idx in self._resolve_indices(keys):
            target = target[idx]
        return target

    def __setitem__(self: T, keys, value):
        indices = self._resolve_indices(keys)
        target = self._data
        for idx in indices[:-1]:
            target = target[idx]
        target[indices[-1]] = value

    def redim(self: T, *new_bounds, preserve=False):
        """
        Resizes the array. 
        If preserve=True, only the upper bound of the last dimension can change.
        """
        if not preserve:
            self.__init__(*new_bounds)
            return

        # VBA ReDim Preserve validation
        if len(new_bounds) != len(self._bounds):
            raise ValueError("Cannot change number of dimensions with Preserve")
        
        for i in range(len(new_bounds) - 1):
            old_l, old_u = self._bounds[i]
            new_l, new_u = new_bounds[i] if isinstance(new_bounds[i], tuple) else (Array.base, new_bounds[i])
            if (old_l, old_u) != (new_l, new_u):
                raise ValueError("Preserve only allows changing the last dimension's upper bound")

        # Capture old data before re-initializing
        old_data_map = {
            indices: self[tuple(l + i for l, i in zip([b[0] for b in self._bounds], indices))]
            for indices in itertools.product(*(range(s) for s in self._shape))
        }

        # Apply new bounds
        self.__init__(*new_bounds)

        # Restore data that fits in new structure
        for rel_indices, value in old_data_map.items():
            try:
                # Map relative 0-based indices to new absolute indices
                abs_indices = tuple(l + i for (l, u), i in zip(self._bounds, rel_indices))
                self[abs_indices] = value
            except IndexError:
                continue # Truncated during resize

    def __repr__(self: T):
        bound_strings = [f"{l} To {u}" for l, u in self._bounds]
        return f"Array({', '.join(bound_strings)})"
