Sub Macro1()
'
' Macro1 Macro
'
'
    
    Dim num As Integer
    Dim i As Integer
    Dim city As String
    
    
    num = Sheets("UK_USA Feeder").Range("Q3").Value
    
    For i = 1 To num
    
    city = Sheets("UK_USA Feeder").Range("P" & i + 1).Value
    
    ActiveWorkbook.SlicerCaches("Slicer_HOTEL_CITY").VisibleSlicerItemsList = Array _
        ( _
        "[UK_USA Feeder].[HOTEL_CITY].&[" & city & "]")
    Sheets("UK_USA Feeder Corr").Select
    Range("T3:Y3").Select
    Selection.Copy
    Sheets("Summary Table").Select
    Range("B" & i + 2).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("UK_USA Feeder Corr").Select
    Range("AI3:AN3").Select
    Selection.Copy
    Sheets("Summary Table").Select
    Range("J" & i + 2).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Sheets("UK_USA Feeder Corr").Select
    Range("AX3:BC3").Select
    Selection.Copy
    Sheets("Summary Table").Select
    Range("R" & i + 2).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    
    
    Sheets("UK_USA Feeder").Select
    Range("M2").Select
    Selection.Copy
    Sheets("Summary Table").Select
    Range("A" & i + 2).Select
    Selection.PasteSpecial xlPasteValues
    
    
    Sheets("UK_USA Feeder").Select
    Range("M2").Select
    Selection.Copy
    Sheets("Summary Table").Select
    Range("I" & i + 2).Select
    Selection.PasteSpecial xlPasteValues
    
    
     Sheets("UK_USA Feeder").Select
    Range("M2").Select
    Selection.Copy
    Sheets("Summary Table").Select
    Range("Q" & i + 2).Select
    Selection.PasteSpecial xlPasteValues
    
    Next i
    
    
    
End Sub
