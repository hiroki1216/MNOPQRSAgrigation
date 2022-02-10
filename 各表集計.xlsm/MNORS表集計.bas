Attribute VB_Name = "MNORS表集計"
Sub MNORS表集計()
Dim ws1 As Worksheet
Dim Worksheet As Worksheet '各ワークシート
Dim worksheetName As String 'ワークシート名用
Dim lastRow As Long 'ワークシート最終行
Dim lastcolumn As Long 'カラム検索用
Dim standardRecord As Long '保険税［料］種別カラムの列番号
Dim i As Long 'カラム検索繰返し処理用
Dim countingIryou As Long '集計結果（医療）
Dim countingKaigo As Long '集計結果（介護)
Dim kensakuRange As range

Set ws1 = ThisWorkbook.Worksheets(1)

For Each Worksheet In Worksheets
'ワークシート名を取得
    worksheetName = Worksheet.Name
    'A列最終行番号の取得
    lastRow = Worksheet.Cells(1, 1).End(xlDown).Row
    
    'NORS表用保険税［料］種別カラムの検索
    lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
    For i = 1 To lastcolumn
        If Worksheet.Cells(1, i).Value = "保険税［料］種別" Then
            standardRecord = i
            Worksheet.Cells(1, standardRecord).Interior.ColorIndex = 3
        End If
    Next i
    
    
    'M表医療分の集計
    If worksheetName Like "*M表*" & "*医療*" Then
        countingIryou = WorksheetFunction.CountA(Worksheet.range("A2:A" & lastRow))
        ws1.range("D7").Value = countingIryou
        
    ElseIf worksheetName Like "*M表*" & "*介護*" Then
        countingKaigo = WorksheetFunction.CountA(Worksheet.range("A2:A" & lastRow))
        ws1.range("G7").Value = countingKaigo
        
    ElseIf worksheetName Like "*N表*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "医療分")
        ws1.range("D8").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "介護分")
        ws1.range("G8").Value = countingKaigo
        
    ElseIf worksheetName Like "*O表*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "医療分")
        ws1.range("D9").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "介護分")
        ws1.range("G9").Value = countingKaigo
        
    ElseIf worksheetName Like "*R表*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "医療分")
        ws1.range("D12").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "介護分")
        ws1.range("G12").Value = countingKaigo

    ElseIf worksheetName Like "*S表*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "医療分")
        ws1.range("D13").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "介護分")
        ws1.range("G13").Value = countingKaigo
    Else
        GoTo Continue
    End If
Continue:
Next

End Sub
