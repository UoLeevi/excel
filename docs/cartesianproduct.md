**CARTESIANPRODUCT** is a custom Excel function that computes the Cartesian product of multiple sets. The function accepts up to nine sets as arguments and returns a two-dimensional array where each row represents a unique combination of elements from the input sets.

**Syntax**

```
=CARTESIANPRODUCT(a_1, [a_2], [a_3], [a_4], [a_5], [a_6], [a_7], [a_8], [a_9])
```

**Arguments**
- `a_1` to `a_9`: The input sets to compute the Cartesian product of.

**Return Value**
- A two-dimensional array where each row represents a unique combination of elements from the input sets.

**Example**

```
=CARTESIANPRODUCT({"A", "B"}, {1, 2, 3}, {"X", "Y"})
```

This example returns a two-dimensional array with six rows and three columns:

```
A   1   X
A   1   Y
A   2   X
A   2   Y
A   3   X
A   3   Y
```

**Notes**
- This function accepts up to nine input sets.
- If any input set is empty, the function returns an empty array.
- If any input set contains non-unique elements, the function returns duplicate rows in the output array.
- This function uses the `MAKEARRAY` and `INDEX` functions, which are only available in Excel 365 or later versions.