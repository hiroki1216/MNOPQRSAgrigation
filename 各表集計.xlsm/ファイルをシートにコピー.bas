Attribute VB_Name = "�t�@�C�����V�[�g�ɃR�s�["
Sub �t�@�C�����V�[�g�ɃR�s�[()

Dim Worksheet As Worksheet '���[�N�V�[�g�擾�p�i�����������j
Dim folderPath As String '�����t�H���_�[�̃f�B���N�g�����擾�p
Dim fileType As String '�t�@�C���̊g���q�p
Dim mergeWorkbook As String '�����t�@�C���p
Dim mergeWorkbookdata As Long '�����t�@�C����A��ŏI�s�擾�p

'�W�v�{�^���𕡐��񉟂��ꂽ�Ƃ��̏����i�V�[�g�̏������j
Application.DisplayAlerts = False
'�h��ی��Ґ��W�v�h�̃��[�N�V�[�g�ȊO���폜
For Each Worksheet In Worksheets
    If Worksheet.Index <> 1 Then
        Worksheet.Delete
    End If
Next
Application.DisplayAlerts = True

'�����t�H���_�[���w��
folderPath = ThisWorkbook.Worksheets(1).range("B2").Value

'�Z���̒l�Ŋg���q���ꍇ����
If Worksheets(1).range("B1").Value = "Excel" Then
    fileType = "\*.xlsx*"
Else
    fileType = "\*.csv"
End If

'�����Ώۃt�@�C�����̎擾
mergeWorkbook = Dir(folderPath & fileType) '(Dir�֐��́A�߂�l�ɁA������^�̃t�@�C������Ԃ�)

'�Ώۃt�@�C���𓝍��p�̃t�@�C���̃V�[�g�ɃR�s�[
Do Until mergeWorkbook = ""
    
    '�}�[�W���郏�[�N�u�b�N�Ƀ}�[�W�����t�@�C�����̃V�[�g��V�K�쐬�h
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = mergeWorkbook
    '�����t�H���_�[\�����t�@�C�������J��
    Workbooks.Open Filename:=folderPath & "\" & mergeWorkbook
    
    '�}�[�W����郏�[�N�u�b�N��A��̍ŏI�s���擾
    mergeWorkbookdata = Workbooks(mergeWorkbook).Worksheets(1).Cells(1, 1).End(xlDown).Row
    
    '�}�[�W����t�@�C���̍ŏI�s�Ƀ}�[�W����郏�[�N�u�b�N���R�s�y
    Workbooks(mergeWorkbook).Worksheets(1).Rows("1:" & mergeWorkbookdata).Copy ThisWorkbook.Worksheets(mergeWorkbook).range("A1")
    Application.DisplayAlerts = False
    Workbooks(mergeWorkbook).Close
    Application.DisplayAlerts = True
    
    '�w��t�H���_�[���̑Ώۃt�@�C�����ċA�I�ɏ���
    mergeWorkbook = Dir()

Loop

End Sub
