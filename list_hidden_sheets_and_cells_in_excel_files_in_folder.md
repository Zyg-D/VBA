```vba
Sub list_hidden_sheets_and_cells_in_excel_files_in_folder()
' Lists into active sheet. Opens and closes all excel files
' Before using - change dir_name
Dim dir_name As String: dir_name = "C:\Temp\test"
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim f As Object
Dim wb As Workbook
Dim ws As Worksheet
Dim c As Range: Set c = ActiveSheet.Cells(1, 1)
Dim r As Range

For Each f In oFSO.GetFolder(dir_name).Files
  If f.Name Like "*.xls*" Then
    Set wb = Application.Workbooks.Open(f)
    For Each ws In wb.Worksheets
      If ws.Visible = xlSheetHidden Then
        c = "Hidden sheet """ & ws.Name & """ in wb """ & wb.Name & """"
        Set c = c.Offset(1)
      ElseIf ws.Visible = xlSheetVeryHidden Then
        c = "Very hidden sheet """ & ws.Name & """ in wb """ & wb.Name & """"
        Set c = c.Offset(1)
      Else
        For Each r In ws.UsedRange.Rows
          If r.Hidden Then
            c = "Hidden row " & r.row & " in sheet """ & ws.Name & """ in wb """ & wb.Name & """"
            Set c = c.Offset(1)
          End If
        Next 'r
        For Each r In ws.UsedRange.Columns
          If r.Hidden Then
            c = "Hidden col " & Split(r.Address, "$")(3) & " in sheet """ & ws.Name & """ in wb """ & wb.Name & """"
            Set c = c.Offset(1)
          End If
        Next 'r
      End If
    Next 'ws
    wb.Close
  End If
Next 'f
End Sub
```
