

```vba
Sub Rows_from_range_sent_to_files_depending_on_designated_column_values()
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim wbActive As Workbook
  Dim ws As Worksheet
  Dim sPathThisFile As String
  Dim sNameThisFile As String
  Dim sPathNewFold As String
  Dim oXLFile As Object
  Set wbActive = ActiveWorkbook
  Set ws = wbActive.ActiveSheet
  Dim lngDataCol As Long
  Dim sAutoFilterAddress As String
  Dim aUnique() As Variant
  Dim ob As Object
  Dim sStulpSkaidymui As String
  Dim aKopinamiStulpai() As Variant
  Dim aKopinamiStulpai2() As Variant
  Dim lFiltPositionModifier As Long
  Dim aAtsirinktosApklausos() As Variant
  Dim sStulpMetaiSem As String
  Dim lStulpMetaiSem As Long
  Dim rColStart As Range
  Dim i As Long
  Dim j As Long
  
  'Stupelis (tikslus range), pagal kuri bus skaidoma i failus:
  sStulpSkaidymui = "C9:C131881"
  
  'Stulpeliai, kurie bus ikeliami i failus:
  aKopinamiStulpai() = Array("C", "D", "F", "L", "AB", "AC")
  
  'Apklausos, kurias imame visais atvejais (stulpelis "metai, sem"):
  aAtsirinktosApklausos() = Array("2015-2016 pavasario", _
                                  "2016-2017 rudens", _
                                  "2016-2017 pavasario", _
                                  "2017-2018 rudens")
  
  'Kuriame stulp yra "metai, sem":
  sStulpMetaiSem = "L"
  '--------------------------------------------------------------------
  
  'irasom filtro adresa
  sAutoFilterAddress = ws.AutoFilter.Range.Address
  
  'filtro pastumimo modifieris
  lFiltPositionModifier = Range(Split(sAutoFilterAddress, ":")(0)).Column - 1
  
  'Kelintame Filtro stulpelyje yra sStulpMetaiSem
  lStulpMetaiSem = ws.Range(sStulpMetaiSem & "1").Column - lFiltPositionModifier
  
  'aKopinamiStulpai2(stulpelio raide, kelintas jis filtre)
  ReDim aKopinamiStulpai2(UBound(aKopinamiStulpai), 1)
  For i = LBound(aKopinamiStulpai) To UBound(aKopinamiStulpai)
    aKopinamiStulpai2(i, 0) = aKopinamiStulpai(i)
    aKopinamiStulpai2(i, 1) = ws.Range(aKopinamiStulpai(i) & "1").Column - lFiltPositionModifier
  Next
  
  'find last data column
  If Application.WorksheetFunction.CountA(ws.Cells) <> 0 Then
    lngDataCol = ws.Cells.Find(What:="*", _
     After:=ws.Range("A1"), _
     Lookat:=xlPart, _
     LookIn:=xlFormulas, _
     SearchOrder:=xlByColumns, _
     SearchDirection:=xlPrevious, _
     MatchCase:=False).Column
  Else
    lngDataCol = 1
  End If

  lngDataCol = lngDataCol + 3 'irasinesim i stulpeli desiniau duomenu
  
  'jeigu netycia butu kas nors is anksto nufiltruota, padarom, kad atsidengtu
  If TestIfFiltered(ws) Then ws.ShowAllData
  'padedam i lapa unique irasus
  ws.Range(sStulpSkaidymui).AdvancedFilter xlFilterCopy, , ws.Cells(1, lngDataCol), True
  'irasom i array unique vertes
  aUnique() = ws.Cells(1, lngDataCol).CurrentRegion.Value
  'istrinam stulpeli, kur irasem unique irasus
  ws.Cells(1, lngDataCol).EntireColumn.Delete
  'uzdedam nuimta filtra ir iskart atfiltruojam tik reikalingas apklausas
  ws.Range(sAutoFilterAddress).AutoFilter lStulpMetaiSem, aAtsirinktosApklausos(), xlFilterValues
  
  'sukuriam nauja fold
  sPathThisFile = wbActive.Path
  sNameThisFile = fso.GetBaseName(wbActive.Name)
  sPathNewFold = sPathThisFile & "\" & sNameThisFile & " - distributed (" & Replace(CStr(Now()), ":", "") & ")"
  fso.CreateFolder (sPathNewFold)
  
  For i = LBound(aUnique) To UBound(aUnique)
    'papildom filtra
    ws.Range(sAutoFilterAddress).AutoFilter _
     Field:=ws.Range(sStulpSkaidymui).Column - lFiltPositionModifier, _
     Criteria1:=aUnique(i, 1)
    
    'Pacheckinam, ar yra rezultatu, ka butu galima kelti i faila
    If (ws.Range(Split(sAutoFilterAddress, ":")(0)).End(xlDown).Row <> ws.Rows.Count) Then
    
      'atidarom nauja wb
      Set oXLFile = Workbooks.Add
      
      'copy and paste filtered values
      For j = LBound(aKopinamiStulpai2) To UBound(aKopinamiStulpai2)
        'ws.AutoFilter.Range.Column(aKopinamiStulpai2(j, 1)).Copy
        'ws.AutoFilter.Range.Range(Cells(1, aKopinamiStulpai2(j, 1)), Cells(ws.AutoFilter.Range.Rows.Count, aKopinamiStulpai2(j, 1))).Copy
        Set rColStart = ws.Range(Split(sAutoFilterAddress, ":")(0)).Offset(0, aKopinamiStulpai2(j, 1) - 1)
        ws.Range(rColStart, ws.Cells(Split(ws.Cells(1, 1).SpecialCells(xlLastCell).Address, "$")(2), rColStart.Column)).Copy
        oXLFile.Sheets(1).Cells(1, j + 1).PasteSpecial xlPasteValuesAndNumberFormats
      Next
      
      oXLFile.SaveAs sPathNewFold & "\" & aUnique(i, 1) & ".xlsx"
      oXLFile.Close
      
      
    End If
    
  Next
  
  'atstatom filtra, kad viska rodytu
  If TestIfFiltered(ws) Then ws.ShowAllData
  
  Set oXLFile = Nothing
  Set fso = Nothing
End Sub


Private Function TestIfFiltered(wsheet As Worksheet) As Boolean
    Dim rngFilter As Range
    Dim r As Long, f As Long
    Set rngFilter = wsheet.AutoFilter.Range
    r = rngFilter.Cells.Count
    f = rngFilter.SpecialCells(xlCellTypeVisible).Count
    If r > f Then TestIfFiltered = True
End Function
```
