Attribute VB_Name = "PQ�\��ی��ҏW�v����"
Sub PQ�\��ی��ҏW�v����()

    Call �����Ώۃt�H���_�[�I��.�����Ώۃt�H���_�[�I��
'�����𑱍s���Ă悢���̊m�F
    Dim rtn As Integer
    rtn = MsgBox("�����t�H���_�[���I������܂����BPQ�\���I���t�H���_�[�ɑ��݂��邱�Ƃ��m�F���Ă��������B" & vbCrLf & "���̂܂܏�����i�߂Ă�낵���ł����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F")
    Select Case rtn
    Case vbYes
        GoTo Continue
    Case vbNo
        Exit Sub
    End Select
Continue:
    Call �t�@�C�����V�[�g�ɃR�s�[.�t�@�C�����V�[�g�ɃR�s�[
    Call PQ�\�p�s�v�f�[�^�̍폜.PQ�\�p�s�v�f�[�^�̍폜
    Call PQ�\�W�v�z��g�p��.PQ�\�W�v�z��g�p��
    MsgBox "PQ�\�̏W�v���������܂����B" & vbCrLf & "�S�Ă̕\�̏W�v���������܂����B" & vbCrLf & "�����ꂳ�܂ł��B", vbInformation
End Sub
