Attribute VB_Name = "MNORS\Wv"
Sub MNORS\Wv()
Dim ws1 As Worksheet
Dim Worksheet As Worksheet 'e[NV[g
Dim worksheetName As String '[NV[gΌp
Dim lastRow As Long '[NV[gΕIs
Dim lastcolumn As Long 'Jυp
Dim standardRecord As Long 'Ϋ―ΕmΏnνΚJΜρΤ
Dim i As Long 'JυJΤ΅p
Dim countingIryou As Long 'WvΚiγΓj
Dim countingKaigo As Long 'WvΚiξμ)
Dim kensakuRange As range

Set ws1 = ThisWorkbook.Worksheets(1)

For Each Worksheet In Worksheets
'[NV[gΌπζΎ
    worksheetName = Worksheet.Name
    'AρΕIsΤΜζΎ
    lastRow = Worksheet.Cells(1, 1).End(xlDown).Row
    
    'NORS\pΫ―ΕmΏnνΚJΜυ
    lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
    For i = 1 To lastcolumn
        If Worksheet.Cells(1, i).Value = "Ϋ―ΕmΏnνΚ" Then
            standardRecord = i
            Worksheet.Cells(1, standardRecord).Interior.ColorIndex = 3
        End If
    Next i
    
    
    'M\γΓͺΜWv
    If worksheetName Like "*M\*" & "*γΓ*" Then
        countingIryou = WorksheetFunction.CountA(Worksheet.range("A2:A" & lastRow))
        ws1.range("D7").Value = countingIryou
        
    ElseIf worksheetName Like "*M\*" & "*ξμ*" Then
        countingKaigo = WorksheetFunction.CountA(Worksheet.range("A2:A" & lastRow))
        ws1.range("G7").Value = countingKaigo
        
    ElseIf worksheetName Like "*N\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "γΓͺ")
        ws1.range("D8").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "ξμͺ")
        ws1.range("G8").Value = countingKaigo
        
    ElseIf worksheetName Like "*O\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "γΓͺ")
        ws1.range("D9").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "ξμͺ")
        ws1.range("G9").Value = countingKaigo
        
    ElseIf worksheetName Like "*R\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "γΓͺ")
        ws1.range("D12").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "ξμͺ")
        ws1.range("G12").Value = countingKaigo

    ElseIf worksheetName Like "*S\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "γΓͺ")
        ws1.range("D13").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "ξμͺ")
        ws1.range("G13").Value = countingKaigo
    Else
        GoTo Continue
    End If
Continue:
Next

End Sub
