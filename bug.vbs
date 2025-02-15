Late Binding and Type Mismatches: VBScript is a loosely typed language, so type mismatches can be difficult to detect. Using late binding (referencing objects without explicit declaration) can lead to runtime errors that are not caught during compilation.  Example: If you expect a function to return a number but it returns a string, you might get unexpected results or a runtime error. 
```vbscript
Dim x
x = "10" + 5
MsgBox x 'Displays "105", not 15
```