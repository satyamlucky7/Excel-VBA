Sub test1()

Application.DisplayAlerts = False

Dim x As Long
Dim y As Long
Dim z As Long
Dim m As String





ThisWorkbook.Sheets("Data Dump").Activate
Range("A2:BK20000").Select
Selection.ClearContents



Workbooks("Master File.xlsb").Worksheets("Pipeline").Activate

ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    
ActiveSheet.ListObjects("T_Pipe").Range.AutoFilter Field:=63, Criteria1:= _
        "=*SAP*", Operator:=xlAnd

Range("A1").Select

Selection.End(xlDown).Select

x = ActiveCell.Row




Range("A1:BL" & x).Select
Selection.Copy

ThisWorkbook.Sheets("Data Dump").Activate
Range("A1").PasteSpecial xlPasteValues

ThisWorkbook.Sheets("Working Sheet").Activate

m = Range("A2").Value


For i = 1 To 1

If m <> "" Then

ThisWorkbook.Sheets("Working Sheet").ShowAllData
Range("A1").Select
Selection.End(xlDown).Select
y = ActiveCell.Row
Range("A2:L" & y).ClearContents
Range("N2:P" & y).ClearContents
Else
End If
Next




ThisWorkbook.Sheets("Data Dump").Activate


Range("A1").Select

Selection.End(xlDown).Select

z = ActiveCell.Row

Range("A2:B" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("A2").PasteSpecial xlPasteValues


ThisWorkbook.Sheets("Data Dump").Activate
Range("E2:F" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("C2").PasteSpecial xlPasteValues


ThisWorkbook.Sheets("Data Dump").Activate
Range("H2:I" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("E2").PasteSpecial xlPasteValues


ThisWorkbook.Sheets("Data Dump").Activate
Range("K2:L" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("G2").PasteSpecial xlPasteValues


ThisWorkbook.Sheets("Data Dump").Activate
Range("O2:P" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("I2").PasteSpecial xlPasteValues



ThisWorkbook.Sheets("Data Dump").Activate
Range("U2:U" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("K2").PasteSpecial xlPasteValues



ThisWorkbook.Sheets("Data Dump").Activate
Range("AE2:AE" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("L2").PasteSpecial xlPasteValues



ThisWorkbook.Sheets("Data Dump").Activate
Range("AF2:AF" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("N2").PasteSpecial xlPasteValues




ThisWorkbook.Sheets("Data Dump").Activate
Range("BK2:BK" & z).Select
Selection.Copy
ThisWorkbook.Sheets("Working Sheet").Activate
Range("P2").PasteSpecial xlPasteValues



Range("N2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.TextToColumns Destination:=Range("N2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
        
Range("M2").Formula = "=COUNTIF($N$2:$N$" & z & " ,N2)"

Range("M2").Copy
Range("M3:M" & z).PasteSpecial xlPasteFormulas





End Sub


