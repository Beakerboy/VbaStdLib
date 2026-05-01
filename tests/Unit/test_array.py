
# Usage Example
Array.base = 1
my_arr = Array((1, 3), (6, 9))
my_arr[1, 6] = "Top Left"
my_arr[3, 9] = "Bottom Right"

# ReDim Preserve my_arr(1 To 3, 6 To 12)
my_arr.redim((1, 3), (6, 12), preserve=True)
print(my_arr[1, 6]) # Output: Top Left
