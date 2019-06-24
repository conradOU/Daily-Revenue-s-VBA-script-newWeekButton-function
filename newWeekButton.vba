Sub newWeekButton()
'
' newWeekButton VBA script
' by Conrad R
'
    Range("A1:I12").Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("B3:H4,B9:H10").ClearContents
    
    Range("B2") = DateAdd("d", 7, Range("B2"))
    Range("C2") = DateAdd("d", 7, Range("C2"))
    Range("D2") = DateAdd("d", 7, Range("D2"))
    Range("E2") = DateAdd("d", 7, Range("E2"))
    Range("F2") = DateAdd("d", 7, Range("F2"))
    Range("G2") = DateAdd("d", 7, Range("G2"))
    Range("H2") = DateAdd("d", 7, Range("H2"))
    
    ActiveSheet.Name = Format(Range("H2"), "Short Date")
    
    ActiveSheet.Previous.Select
    Range("I14").Copy
    ActiveSheet.Next.Select
    Range("I14").Select
    ActiveSheet.Paste
    
End Sub
