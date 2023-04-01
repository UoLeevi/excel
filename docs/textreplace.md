**TEXTREPLACE** is a custom Excel function that replaces multiple instances of text in a string with corresponding replacement text. The function accepts a string and an array of alternating old and new text pairs as arguments, and returns the updated string with all occurrences of old text replaced with new text.

**Syntax**
```
=TEXTREPLACE(text, old_text_new_text_alternating_array)`
```

**Arguments**

- `text`: Required. The text string to search and replace text in.
' `old_text_new_text_alternating_array`: Required. A one-dimensional array of alternating old and new text values to replace in the input text string.

**Return Value**
- The updated text string with all occurrences of old text replaced with new text.

**Example**
```
=TEXTREPLACE("Hello, World!", {"Hello", "Hi", "World", "Universe"})
```

This example returns the updated text string:

```
"Hi, Universe!"
```

**Notes**
- This function replaces all occurrences of old text in the input text string with corresponding new text.
- If the input text string contains no instances of old text, the function returns the original text string.
- This function is case-sensitive.
- This function uses the SEQUENCE, INDEX, and REDUCE functions, which are only available in Excel 365 or later versions.