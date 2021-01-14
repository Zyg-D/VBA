

```vba
Sub unmerge_keeping_values_in_all_cells()
' Unmerges all cells in wb, all the new cells get the value of the merged cell
  Dim ws As Worksheet
  Dim c As Range
  Dim cc As Range
  Dim rngMerged As Range
  Dim x As Variant
  
  For Each ws In ThisWorkbook.Sheets
    For Each c In ws.UsedRange
      If c.MergeCells Then
        Set rngMerged = c.MergeArea
        x = rngMerged.Cells(1, 1).Value
        rngMerged.UnMerge
        For Each cc In rngMerged
          cc.Value = x
        Next
      End If
    Next
  Next
    
End Sub
```
