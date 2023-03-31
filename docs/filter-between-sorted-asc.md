**FILTER_BETWEEN_SORTED_ASC** is a custom Excel function that filters a range of cells based on a set of criteria. The function assumes that the criteria range is sorted in ascending order.

**Syntax**
```
=FILTER_BETWEEN_SORTED_ASC(range, sorted_criteria_range, criteria_from, criteria_to)
```


**Arguments**
- `range`: Required. The range of cells to filter.
- `sorted_criteria_range`: Required. The range of cells that contains the criteria values, which is assumed to be sorted in ascending order.
- `criteria_from`: Required. The lower boundary of the filter range.
- `criteria_to`: Required. The upper boundary of the filter range.

**Return Value**
- A range of cells that match the specified criteria and are located within the `range`.

**Example**
```
=FILTER_BETWEEN_SORTED_ASC(A2:A100, B2:B100, 10, 20)
```


This example returns a range of cells within the `A2:A100` range that have corresponding values in the `B2:B100` range between `10` and `20`, inclusive.

**Notes**
- This function uses the `XMATCH` and `INDEX` functions, which are only available in Excel 365 or later versions.
- If the `sorted_criteria_range` is not sorted in ascending order, the results of this function may be unexpected.
