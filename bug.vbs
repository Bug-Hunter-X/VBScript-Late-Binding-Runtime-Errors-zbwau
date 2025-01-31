Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where version compatibility isn't guaranteed.  For example, attempting to use a method that's been removed from a newer version of an object will cause a runtime error.

Example:
```vbscript
Set objExcel = CreateObject("Excel.Application")
'Assume a method 'MyMethod' is removed in a newer Excel version
 objExcel.MyMethod
```
This may work fine with one version of Excel but fail with another. Early binding (declaring object types explicitly) would be preferable but requires explicit references in your project.
