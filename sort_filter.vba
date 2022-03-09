Sub eMAIL_FILTER()

    Cells.Replace What:=", ", Replacement:=" ", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, FormulaVersion:=xlReplaceFormula2
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "New INTTRA"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-1],4) = ""OOLU"", ""xxx"",IF(LEFT(RC[-1],4)=""ONEY"", ""xxx"",IF(LEFT(RC[-1],4)=""ZIMU"", ""xxx"",IF(LEFT(RC[-1],4)=""MAEU"", ""xxx"",IF(LEFT(RC[-1],4)="""",""xxx"",RC[-1])))))"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & Range("A" & Rows.Count).End(xlUp).Row), Type:=xlFillDefault
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Range(Selection, Selection.End(xlDown)).PasteSpecial Paste:=xlPasteValues
    
    Cells.Replace What:=",", Replacement:=" ", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, FormulaVersion:=xlReplaceFormula2
    
    Columns("I:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    
End Sub
Sub sort_data()

    Call sort_emailFilter.eMAIL_FILTER
    
     Dim finalrow As Long
     Dim i As Long
     finalrow = Cells(Rows.Count, 2).End(xlUp).Row
     For i = finalrow To 2 Step -1
        If IsEmpty(Cells(i, 6)) Then Cells(i, 6).Delete Shift:=xlUp
     Next i
     
     For i = finalrow To 2 Step -1
        If Cells(i, 7) = "xxx" Then Cells(i, 7).Delete Shift:=xlUp
    Next i
    
    For i = finalrow To 2 Step -1
        If IsEmpty(Cells(i, 8)) Then Cells(i, 8).Delete Shift:=xlUp
    Next i
End Sub
