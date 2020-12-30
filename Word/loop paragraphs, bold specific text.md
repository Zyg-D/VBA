```vba
Sub Bold_ending_parentheses()
' Loops through paragraphs and finds text in parentheses which ends the paragraph.
' Then the text is bolded.
  Dim par As Word.Paragraph
  Dim str As String
  Dim closes As Byte
  Dim opens As Long

  For Each par In ActiveDocument.Paragraphs
    str = StrReverse(par.Range.Text)
    closes = InStr(Left(str, 4), ")")
    If closes Then
      opens = InStr(str, "(")
      If opens Then
        
        With par.Range
          .Find.Text = StrReverse(Mid(str, closes, opens - closes + 1))
          Do
            .Find.Execute
            If .Find.Found Then .Font.Bold = True
          Loop While .Find.Found
        End With
        
      End If
    End If
  Next
End Sub
```
