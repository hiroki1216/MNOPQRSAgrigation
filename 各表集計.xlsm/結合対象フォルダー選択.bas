Attribute VB_Name = "結合対象フォルダー選択"
Sub 結合対象フォルダー選択()
    '処理用のフォルダーを取得
    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        range("b2").Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If
End Sub
