Attribute VB_Name = "MNORS�\��ی��ҏW�v����"
Sub MNORS�\��ی��Ґ��̏W�v����()
    Call �����Ώۃt�H���_�[�I��.�����Ώۃt�H���_�[�I��
    '�����𑱍s���Ă悢���̊m�F
    Dim rtn As Integer
    rtn = MsgBox("�����t�H���_�[���I������܂����B" & vbCrLf & "���̂܂܏�����i�߂Ă�낵���ł����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F")
    Select Case rtn
    Case vbYes
        GoTo Continue
    Case vbNo
        Exit Sub
    End Select
Continue:
    Call �t�@�C�����V�[�g�ɃR�s�[.�t�@�C�����V�[�g�ɃR�s�[
    Call MNORS�\�p�s�v�f�[�^�̍폜.MNORS�\�p�s�v�f�[�^�̍폜
    Call MNORS�\�W�v.MNORS�\�W�v
    ThisWorkbook.Worksheets(1).Select
    MsgBox "MNORS�\�̏W�v���������܂����B" & vbCrLf & "������PQ�\�W�v�{�^�����N���b�N���Ă��������B", vbInformation
End Sub
