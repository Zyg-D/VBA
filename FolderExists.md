performance not tested, the first function is the 
```vba 
Public Function DirExists(sDir As String) As Boolean
    Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    DirExists = oFSO.FolderExists(sDir)
End Function

' On Error Resume Next, because of these cases: "AA"; "::"; "//"; "\\"; ""
Public Function DirExists2(sDir As String) As Boolean
    On Error Resume Next
    DirExists2 = ((GetAttr(sDir) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

' This also changes the current directory, which is not really used, so np
' On Error Resume Next, because of these cases: "AA"; "::"; "//"; "\\"; ""
Public Function DirExists3(sDir As String) As Boolean
  Err.Clear
  On Error Resume Next
  ChDir (sDir)
  DirExists3 = (Err.Number = 0)
  On Error GoTo 0
End Function

' This function is bad, because it returns True when "" is passed
' On Error Resume Next, because of these cases: "//"; "\\"
Public Function DirExists4(sDir As String) As Boolean
  On Error Resume Next
  DirExists4 = Len(Dir(sDir, vbDirectory)) > 0
  On Error GoTo 0
End Function

' This function is bad, because it returns True when "" is passed
' On Error Resume Next, because of these cases: "//"; "\\"
Public Function DirExists5(sDir As String) As Boolean
  'On Error Resume Next
  DirExists5 = Dir(sDir, vbDirectory) <> ""
  On Error GoTo 0
End Function
```
Testing function:
```vba
Sub tst()
'  test cases:
'  "C:"
'  "C:/"
'  "C:\"
'  "C:/Temp"
'  "C:\Temp"
'  "C:/Temp/"
'  "C:\Temp\"
'  "::"
'  "//"
'  "\\"
'  ""
'  "AA"
'  "C:/Temp" & ChrW(261) 'ChrW(261)=Letter A with ogonek
'  "C:\Temp" & ChrW(261) 'ChrW(261)=Letter A with ogonek
'  "C:/Temp" & ChrW(261) & "/" 'ChrW(261)=Letter A with ogonek
'  "C:\Temp" & ChrW(261) & "\" 'ChrW(261)=Letter A with ogonek

  Dim s As String: s = "AA"
  Debug.Print "1 " & DirExists(s) & vbTab _
            & "2 " & DirExists2(s) & vbTab _
            & "3 " & DirExists3(s) & vbTab _
            & "4 " & DirExists4(s) & vbTab _
            & "5 " & DirExists5(s)
End Sub
```
