Creates a folder together with parent folders, if needed. 
"Waits" until `cmd` creates it (time-out at 2 seconds). 

```vba
Public Function DirCreate(sPath As String)
  Dim timedOut As Boolean
  Dim t As Date: t = Now + TimeValue("0:00:02")
  If Not DirExists(sPath) Then Shell ("cmd /c md """ & sPath & """")
  Do: timedOut = Now > t
  Loop Until DirExists(sPath) Or timedOut
  If timedOut Then Err.Raise vbObjectError + 513, , "The folder was not created"
End Function

Public Function DirExists(sDir As String) As Boolean
  Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
  DirExists = oFSO.FolderExists(sDir)
End Function
```

