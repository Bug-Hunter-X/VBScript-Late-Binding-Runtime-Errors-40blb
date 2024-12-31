Improved error handling is crucial to prevent unexpected crashes. Here's a more robust version:

```vbscript
On Error Resume Next

Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
  MsgBox "Error creating Excel object: " & Err.Description, vbCritical
  WScript.Quit
End If

Set objWorkbook = objExcel.Workbooks.Add

' Add error handling
Err.Clear
On Error GoTo SaveError

Call objWorkbook.SaveAs("C:\myFile.xlsx", 51) 

On Error GoTo 0

' Clean up
objWorkbook.Close
Set objWorkbook = Nothing
objExcel.Quit
Set objExcel = Nothing

Exit Sub

SaveError:
MsgBox "Error saving file: " & Err.Description, vbCritical

' Clean up even on error
On Error Resume Next
If Not objWorkbook Is Nothing Then objWorkbook.Close
If Not objExcel Is Nothing Then objExcel.Quit
On Error GoTo 0

WScript.Quit
```

This enhanced code uses error handling (`On Error Resume Next`, `On Error GoTo`, `Err.Number`, `Err.Description`) to catch potential errors during object creation and file saving.  It also includes cleanup code within the error handler to ensure resources are released even if errors occur.