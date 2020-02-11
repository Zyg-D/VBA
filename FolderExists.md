performance not tested, the first function is the 
```vba 
Public Function DirExists(sDir As String) As Boolean
    Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    DirExists = oFSO.FolderExists(sDir)
End Function

Public Function DirExists2(sDir As String) As Boolean
    On Error Resume Next
    DirExists2 = ((GetAttr(sDir) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

' This also changes the current directory, which is not really used, so np
Public Function DirExists3(sDir As String) As Boolean
  Err.Clear
  On Error Resume Next
  ChDir (sDir)
  DirExists3 = (Err.Number = 0)
  On Error GoTo 0
End Function
```
Testing function:
```vba
Sub tst()
  Dim s As String: s = "c:/Temp"
  Debug.Print "1 " & DirExists(s) & vbTab _
            & "2 " & DirExists2(s) & vbTab _
            & "3 " & DirExists3(s)
End Sub
```
Testing cases:
```
C:
C:/
C:\
C:/Temp
C:\Temp
C:/Temp/
C:\Temp\
::
//
\\
AAA
```
