Attribute VB_Name = "�����Ώۃt�H���_�[�I��"
Sub �����Ώۃt�H���_�[�I��()
    '�����p�̃t�H���_�[���擾
    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        range("b2").Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If
End Sub
