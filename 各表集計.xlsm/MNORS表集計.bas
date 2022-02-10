Attribute VB_Name = "MNORS�\�W�v"
Sub MNORS�\�W�v()
Dim ws1 As Worksheet
Dim Worksheet As Worksheet '�e���[�N�V�[�g
Dim worksheetName As String '���[�N�V�[�g���p
Dim lastRow As Long '���[�N�V�[�g�ŏI�s
Dim lastcolumn As Long '�J���������p
Dim standardRecord As Long '�ی��Łm���n��ʃJ�����̗�ԍ�
Dim i As Long '�J���������J�Ԃ������p
Dim countingIryou As Long '�W�v���ʁi��Áj
Dim countingKaigo As Long '�W�v���ʁi���)
Dim kensakuRange As range

Set ws1 = ThisWorkbook.Worksheets(1)

For Each Worksheet In Worksheets
'���[�N�V�[�g�����擾
    worksheetName = Worksheet.Name
    'A��ŏI�s�ԍ��̎擾
    lastRow = Worksheet.Cells(1, 1).End(xlDown).Row
    
    'NORS�\�p�ی��Łm���n��ʃJ�����̌���
    lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
    For i = 1 To lastcolumn
        If Worksheet.Cells(1, i).Value = "�ی��Łm���n���" Then
            standardRecord = i
            Worksheet.Cells(1, standardRecord).Interior.ColorIndex = 3
        End If
    Next i
    
    
    'M�\��Õ��̏W�v
    If worksheetName Like "*M�\*" & "*���*" Then
        countingIryou = WorksheetFunction.CountA(Worksheet.range("A2:A" & lastRow))
        ws1.range("D7").Value = countingIryou
        
    ElseIf worksheetName Like "*M�\*" & "*���*" Then
        countingKaigo = WorksheetFunction.CountA(Worksheet.range("A2:A" & lastRow))
        ws1.range("G7").Value = countingKaigo
        
    ElseIf worksheetName Like "*N�\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "��Õ�")
        ws1.range("D8").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "��앪")
        ws1.range("G8").Value = countingKaigo
        
    ElseIf worksheetName Like "*O�\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "��Õ�")
        ws1.range("D9").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "��앪")
        ws1.range("G9").Value = countingKaigo
        
    ElseIf worksheetName Like "*R�\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "��Õ�")
        ws1.range("D12").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "��앪")
        ws1.range("G12").Value = countingKaigo

    ElseIf worksheetName Like "*S�\*" Then
        Worksheet.Select
        Set kensakuRange = Worksheet.range(Cells(2, standardRecord), Cells(lastRow, standardRecord))
        countingIryou = WorksheetFunction.CountIf(kensakuRange, "��Õ�")
        ws1.range("D13").Value = countingIryou
        countingKaigo = WorksheetFunction.CountIf(kensakuRange, "��앪")
        ws1.range("G13").Value = countingKaigo
    Else
        GoTo Continue
    End If
Continue:
Next

End Sub
