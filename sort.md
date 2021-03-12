```vba
Sub sort()
  Dim sList  As String
  Dim rs As Object: Set rs = CreateObject("ADODB.RECORDSET")
  Dim i As Long
  Dim splitted As Variant
  
  splitted = Split("B, A, C", ", ")
  
  ' Put data into rs
  rs.CursorType = 3 '3=adOpenStatic
  rs.Fields.append "s", 200, 25 '200=adVarChar
  rs.Open
  For i = 0 To UBound(splitted)
    rs.AddNew "s", splitted(i)
    rs.Update
  Next
  
  rs.sort = "s asc" 'ascending
  
  ' Read data from rs
  rs.MoveFirst
  Do Until rs.EOF
    sList = sList & vbCrLf & rs.Fields("s")
    rs.MoveNext
  Loop
  
  MsgBox sList

End Sub
```
