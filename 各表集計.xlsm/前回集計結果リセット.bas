Attribute VB_Name = "�O��W�v���ʃ��Z�b�g"
Sub �O��W�v���ʃ��Z�b�g()
'�����𑱍s���Ă悢���̊m�F
    Dim rtn As Integer
    rtn = MsgBox("�O�񌋉ʂ��폜���܂��B" & vbCrLf & "��낵���ł����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F")
    Select Case rtn
    Case vbYes
        GoTo Continue
    Case vbNo
        Exit Sub
    End Select
Continue:
    '�O��W�v���ʂ̒l���N���A
    
    'MNORS�\�̒l�N���A����
    ThisWorkbook.Worksheets(1).range("D7:D9").ClearContents
    ThisWorkbook.Worksheets(1).range("D12:D13").ClearContents
    ThisWorkbook.Worksheets(1).range("G7:G9").ClearContents
    ThisWorkbook.Worksheets(1).range("G12:G13").ClearContents
    'PQ�\�̒l�N���A����
    ThisWorkbook.Worksheets(1).range("D20:D21").ClearContents
    ThisWorkbook.Worksheets(1).range("D23:D24").ClearContents
    ThisWorkbook.Worksheets(1).range("G20:G21").ClearContents
    ThisWorkbook.Worksheets(1).range("G23:G24").ClearContents
       
    '�V�[�g�̏�����
    Application.DisplayAlerts = False
    '�h��ی��Ґ��W�v�h�̃��[�N�V�[�g�ȊO���폜
    For Each Worksheet In Worksheets
        If Worksheet.Index <> 1 Then
            Worksheet.Delete
        End If
    Next
    Application.DisplayAlerts = True
    
    MsgBox "�O��W�v���ʂ̃��Z�b�g���������܂����B", vbInformation
End Sub
