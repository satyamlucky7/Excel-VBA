Sub remove10count()


Dim x As Long
Dim y As Long
Dim z As Long
Dim a As String
Dim b As Long
Dim c As Long
Dim d As String
Dim n As Long

Dim lastrow As Long
Dim count As Long
Dim outcount As Long


ThisWorkbook.Sheets("Working Sheet").Activate



Range("A1").Select

Selection.End(xlDown).Select

x = ActiveCell.Row



ActiveSheet.Range("$A$1:$P$" & x).AutoFilter Field:=13, Criteria1:="10"

ActiveSheet.Range("$A$1:$P$" & x).AutoFilter Field:=11, Criteria1:="<0", _
        Operator:=xlAnd

Range("K1").Select


Selection.End(xlDown).Select

y = ActiveCell.Row


Range("N" & y).Select

Range(Selection, Selection.End(xlUp)).Select

Selection.Copy

ThisWorkbook.Sheets("Backup Sheet").Activate

Range("G1").Select

Selection.PasteSpecial xlPasteValues

ActiveSheet.Range("$G$1:$G$1000").RemoveDuplicates Columns:=1, Header:=xlYes
Rows("1:1").Select
Selection.Delete Shift:=xlUp



Range("G1").Select

Selection.End(xlDown).Select

z = 1


ThisWorkbook.Sheets("Working Sheet").Activate
ActiveSheet.Range("$A$1:$P$" & x).AutoFilter Field:=11

For i = 1 To 1
ThisWorkbook.Sheets("Backup Sheet").Activate

d = Range("G" & i).Value

ThisWorkbook.Sheets("Working Sheet").Activate

ActiveSheet.Range("$A$1:$P$" & x).AutoFilter Field:=14, Criteria1:=d

'------------done till here ------------------'

 
 Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate


Range("A1").Select
Selection.End(xlDown).Select
lastrow = ActiveCell.Row


Range("T1").Formula = "=SUBTOTAL(3,Q1:Q" & lastrow & ")"
Range("U1").Formula = "=SUBTOTAL(3,R1:R" & lastrow & ")"
count = Range("T1").Value - 1



If count = 4 Then


 Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

 Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete


 Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

outcount = Range("U1").Value - 1

If outcount = 0 Then

Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete
Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete
Columns("N:N").Select


ElseIf outcount = 1 Then

Columns("E:E").Select
    Selection.Find(What:="Outbound", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete
Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

ElseIf outcount = 2 Then

Columns("E:E").Select
    Selection.Find(What:="Outbound", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete
Columns("E:E").Select
    Selection.Find(What:="Outbound", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete



Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete


ElseIf outcount = 5 Then

Columns("E:E").Select
    Selection.Find(What:="Outbound", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete
Columns("E:E").Select
    Selection.Find(What:="Outbound", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete

Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
b = ActiveCell.Row

Range(b & ":" & b).Delete


Range("A1").Select


Selection.End(xlDown).Select

b = ActiveCell.Row

Range(b & ":" & b).Delete


End If

ElseIf count = 1 Then


 Columns("N:N").Select
    Selection.Find(What:=d, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
      
       

b = ActiveCell.Row

Range(b & ":" & b).Delete

Range("A1").Select


Selection.End(xlDown).Select

b = ActiveCell.Row

Range(b & ":" & b).Delete


End If

    ActiveSheet.Range("$A$1:$P$" & x).AutoFilter Field:=14
Next




ActiveSheet.Range("$A$1:$P$" & x).AutoFilter Field:=13



End Sub
