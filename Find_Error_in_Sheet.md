```vba
Sub Find_Error_In_Sheet()
' Finds and selects the first error value in a sheet.

  Dim rUsed As Range: Set rUsed = ThisWorkbook.ActiveSheet.UsedRange
  Dim rConstErrs As Range
  Dim rFormErrs As Range
  
  On Error Resume Next
  
  ' Looking for errors in simple values
  Set rConstErrs = rUsed.SpecialCells(xlCellTypeConstants, xlErrors)
  If Not (Err = 1004 Or Err = 0) Then
    MsgBox "Error setting rConstErrs! " & Err & " " & Err.Description
    Exit Sub
  End If
  
  ' Looking for errors resulting from formula expressions
  Set rFormErrs = rUsed.SpecialCells(xlCellTypeFormulas, xlErrors)
  If Not (Err = 1004 Or Err = 0) Then
    MsgBox "Error setting rFormErrs! " & Err & " " & Err.Description
    Exit Sub
  End If
  
  Err = 0
  On Error GoTo 0
  
  If Not rConstErrs Is Nothing Then
    rConstErrs.Cells(1, 1).Select
  ElseIf Not rFormErrs Is Nothing Then
    rFormErrs.Cells(1, 1).Select
  Else
    MsgBox "Success! No cells containing error values were found " _
    & "in the UsedRange - " & rUsed.Address
  End If
  
End Sub
```
