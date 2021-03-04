```vba
Sub list_all_re_matches()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  Dim re_matches As Object, m As Object
  With re
    .Pattern = "He(l+)o"
    .Global = 1 '1 - finds all matches, 0 - only the 1st
    .IgnoreCase = 0
    .MultiLine = 1 'pattern matching happens across line breaks
  End With
  Set re_matches = re.Execute("Helo Hello Hellllo Heo World!")
  For Each m In re_matches
    Debug.Print m.Value 'Result: Helo Hello Hellllo
  Next
End Sub
```

```vba
Sub re_replace()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  With re
    .Pattern = "hello"
    .Global = 1 '1 - finds all matches, 0 - only the 1st
    .IgnoreCase = 1
    .MultiLine = 1 'pattern matching happens across line breaks
  End With
  Debug.Print re.Replace("Hello hello World!", "Bye")
End Sub
```

```vba
Sub re_match_exists()
  Dim re As Object: Set re = CreateObject("vbscript.regexp")
  With re
    .Pattern = "hello"
    .Global = 1 '1 - finds all matches, 0 - only the 1st
    .IgnoreCase = 1
    .MultiLine = 1 'pattern matching happens across line breaks
  End With
  Debug.Print re.Test("Hello World!")
End Sub
```
