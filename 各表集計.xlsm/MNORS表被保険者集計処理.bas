Attribute VB_Name = "MNORS表被保険者集計処理"
Sub MNORS表被保険者数の集計処理()
    Call 結合対象フォルダー選択.結合対象フォルダー選択
    '処理を続行してよいかの確認
    Dim rtn As Integer
    rtn = MsgBox("処理フォルダーが選択されました。" & vbCrLf & "このまま処理を進めてよろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
    Select Case rtn
    Case vbYes
        GoTo Continue
    Case vbNo
        Exit Sub
    End Select
Continue:
    Call ファイルをシートにコピー.ファイルをシートにコピー
    Call MNORS表用不要データの削除.MNORS表用不要データの削除
    Call MNORS表集計.MNORS表集計
    ThisWorkbook.Worksheets(1).Select
    MsgBox "MNORS表の集計が完了しました。" & vbCrLf & "続いてPQ表集計ボタンをクリックしてください。", vbInformation
End Sub
