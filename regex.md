```vba
Sub re_match_exists()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  With re
    .Pattern = "hello"
    .Global = 1 '1 = finds all matches, 0 (def) = only the 1st
    .IgnoreCase = 1 'def = 0
    .MultiLine = 1 '1 = ^$ surrounds lines instead of the whole str; def = 0
  End With
  Debug.Print re.Test("Hello World!")
End Sub
```

```vba
Sub list_full_re_matches()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  Dim re_matches As Object, m As Object
  With re
    .Pattern = "He(l+)o"
    .Global = 1 '1 = finds all matches, 0 (def) = only the 1st
    .IgnoreCase = 0 'def = 0
    .MultiLine = 1 '1 = ^$ surrounds lines instead of the whole str; def = 0
  End With
  Set re_matches = re.Execute("Helo Hello Hellllo Heo World!")
  For Each m In re_matches
    Debug.Print m.Value 'Result: Helo Hello Hellllo
  Next
End Sub
```

```vba
Sub list_groups()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  Dim full_p_matches As Object, m As Object
  Dim i As Long
  With re
    .Pattern = "(\d+) (\w+)"
    .Global = 1 '1 = finds all matches, 0 (def) = only the 1st
    .IgnoreCase = 0 'def = 0
    .MultiLine = 1 '1 = ^$ surrounds lines instead of the whole str; def = 0
  End With
  Set full_p_matches = re.Execute("2020 Mar")
  For Each m In full_p_matches
    For i = 0 To m.submatches.Count - 1
      Debug.Print m.submatches.Item(i)
    Next
  Next
End Sub
```

```vba
Sub re_replace()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  With re
    .Pattern = "hello"
    .Global = 1 '1 = finds all matches, 0 (def) = only the 1st
    .IgnoreCase = 1 'def = 0
    .MultiLine = 1 '1 = ^$ surrounds lines instead of the whole str; def = 0
  End With
  Debug.Print re.Replace("Hello hello World!", "Bye")
End Sub
```
