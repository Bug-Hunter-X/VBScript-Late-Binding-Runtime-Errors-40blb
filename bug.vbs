Late Binding in VBScript can lead to runtime errors if an object or method doesn't exist.  Consider this code:

```vbscript
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add

' This line might fail if the method doesn't exist
Call objWorkbook.SaveAs("C:\myFile.xlsx", 51) '51 is for xlsx format
```

If the Excel object doesn't have a `SaveAs` method with the specified arguments or the file path is incorrect, this will throw a runtime error that might not be easily caught during development.