```vba
Sub calling_func()
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
Debug.Print GivePath("c1.xlsx", oFSO.GetFolder("C:\Temp"))
End Sub

Private Function GivePath(sName As String, oDir As Object, Optional ByRef blnFound As Boolean) As String
  Dim f As Object
  Dim sf As Object
  
  If blnFound Then Exit Function
  
  For Each f In oDir.Files
    If f.Name = sName Then
      GivePath = f.Path
      blnFound = True
    End If
    If blnFound Then Exit Function
  Next
  
  For Each sf In oDir.SubFolders
    GivePath = GivePath(sName, sf, blnFound)
    If blnFound Then Exit Function
  Next
  
  If GivePath = "" Then GivePath = "The file was not found"
  
End Function
```
