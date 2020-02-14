Not tested for performance.

**Opt.1.** Creates a folder together with parent folders, if needed. 
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

**Opt.2.** Recursive VBA. 

```vba
Function MakeDir(ByVal sDir As String)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(sDir) Then
    MakeDir (fso.GetParentFolderName(sDir))
    fso.CreateFolder sDir
  End If
End Function
```

