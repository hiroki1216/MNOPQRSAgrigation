Attribute VB_Name = "PQ表被保険者集計処理"
Sub PQ表被保険者集計処理()

    Call 結合対象フォルダー選択.結合対象フォルダー選択
'処理を続行してよいかの確認
    Dim rtn As Integer
    rtn = MsgBox("処理フォルダーが選択されました。PQ表が選択フォルダーに存在することを確認してください。" & vbCrLf & "このまま処理を進めてよろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
    Select Case rtn
    Case vbYes
        GoTo Continue
    Case vbNo
        Exit Sub
    End Select
Continue:
    Call ファイルをシートにコピー.ファイルをシートにコピー
    Call PQ表用不要データの削除.PQ表用不要データの削除
    Call PQ表集計配列使用版.PQ表集計配列使用版
    MsgBox "PQ表の集計が完了しました。" & vbCrLf & "全ての表の集計が完了しました。" & vbCrLf & "おつかれさまです。", vbInformation
End Sub
