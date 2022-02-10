Attribute VB_Name = "����t�H���_�[���t�@�C��������"
Sub ����t�H���_�[���t�@�C��������()
On Error Resume Next

'�W�v�{�^���𕡐��񉟂��ꂽ�Ƃ��̏����i�V�[�g�̏������j
Application.DisplayAlerts = False
Worksheets("merge").Delete
Application.DisplayAlerts = True

'�hmerge�t�@�C�����쐬�h
Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "merge"

Dim folderPath
folderPath = ThisWorkbook.Worksheets(1).range("B2").Value '�����t�H���_�[���w��

'�g���q�̎�ނŏꍇ����
Dim fileType
If ThisWorkbook.Worksheets(1).range("B1").Value = "Excel" Then
    fileType = "\*.xlsx*"
Else
    fileType = "\*.csv"
End If

Dim mergeWorkbook
'��������t�@�C��(excel�ł����Ƃ����book)���擾
mergeWorkbook = Dir(folderPath & fileType) '�����t�@�C�������擾(Dir�֐��́A�߂�l�Ƃ��ăt�@�C������Ԃ�)

'�ŏ��̍s���R�s��
 Workbooks.Open Filename:=folderPath & "\" & mergeWorkbook '�����t�H���_�[\�����t�@�C�������J��
 Workbooks(mergeWorkbook).Worksheets(1).Rows(1).Copy ThisWorkbook.Worksheets("merge").range("A1")
 Application.DisplayAlerts = False
 Workbooks(mergeWorkbook).Close
 Application.DisplayAlerts = True
 
 '�f�[�^���ʗp�̃t�@�C�����J�������ŏI��ɒǉ�
 Dim lastcolumn As Long
 lastcolumn = ThisWorkbook.Worksheets("merge").Cells(1, Columns.Count).End(xlToLeft).Column
 ThisWorkbook.Worksheets("merge").Cells(1, lastcolumn + 1).Value = "�t�@�C����"
 
 Dim counter As Long
 counter = 0
 

'�������ʂ������Ȃ�܂ŌJ�ԏ���
Do Until mergeWorkbook = ""
    Workbooks.Open Filename:=folderPath & "\" & mergeWorkbook '�����t�H���_�[\�����t�@�C�������J��
    If Workbooks.Count = 20 Then Exit Do

    Dim mergeWorkbookdata
    Dim thisWorkbookdata
    Dim i

    '�J�����G�N�Z���t�@�C���ɕ����̃V�[�g�����݂���ꍇ�́A�V�[�g���A�J�ԏ������s��
    For i = 1 To Workbooks(mergeWorkbook).Worksheets.Count

    '�}�[�W����郏�[�N�u�b�N��A��̍ŏI�s
    mergeWorkbookdata = Workbooks(mergeWorkbook).Worksheets(i).range("a" & Rows.Count).End(xlUp).Row
    '�}�[�W���郏�[�N�u�b�N��A��̍ŏI�s
    thisWorkbookdata = ThisWorkbook.Worksheets("merge").range("a" & Rows.Count).End(xlUp).Row
    
    '�}�[�W����t�@�C���̍ŏI�s�̂ɑ����āA�}�[�W����郏�[�N�u�b�N���R�s�y
    Workbooks(mergeWorkbook).Worksheets(i).Rows("2:" & mergeWorkbookdata).Copy ThisWorkbook.Worksheets("merge").range("a" & thisWorkbookdata + 1)
    
    
    Next i
    
     '�t�@�C������ǉ�
    Dim lastRow As Long
    Dim startrow As Long
    Dim s As Long
    
    'A��ŏI�s�ԍ�
    lastRow = ThisWorkbook.Worksheets("merge").Cells(1, 1).End(xlDown).Row
    '�ŏI��ԍ�
    lastcolumn = ThisWorkbook.Worksheets("merge").Cells(1, Columns.Count).End(xlToLeft).Column
    '�ŏI����͍ςݍŏI�s�ԍ�
    
    '�t�@�C�����o�͏����𓝍��񐔂ŏꍇ�����i�P��ڂ̂ݏo�͂��Ȃ����߁w�Q�x������j
    If counter = 0 Then
        startrow = ThisWorkbook.Worksheets("merge").Cells(1, 1).Row + 1
    Else
        startrow = ThisWorkbook.Worksheets("merge").Cells(1, lastcolumn).End(xlDown).Row + 1
    End If
    
    
    Debug.Print (counter)
    Debug.Print (lastRow)
    Debug.Print (lastcolumn)
    Debug.Print (startrow)
    
    For s = startrow To lastRow
        ThisWorkbook.Worksheets("merge").Cells(s, lastcolumn).Value = mergeWorkbook
    Next s

    Application.DisplayAlerts = False
    Workbooks(mergeWorkbook).Close
    Application.DisplayAlerts = True

    mergeWorkbook = Dir()
    
    '�t�@�C�����������ꂽ�񐔕����Z����
    counter = counter + 1
Loop



End Sub

