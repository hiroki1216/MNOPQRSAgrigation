Attribute VB_Name = "前回集計結果リセット"
Sub 前回集計結果リセット()
'処理を続行してよいかの確認
    Dim rtn As Integer
    rtn = MsgBox("前回結果を削除します。" & vbCrLf & "よろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
    Select Case rtn
    Case vbYes
        GoTo Continue
    Case vbNo
        Exit Sub
    End Select
Continue:
    '前回集計結果の値をクリア
    
    'MNORS表の値クリア処理
    ThisWorkbook.Worksheets(1).range("D7:D9").ClearContents
    ThisWorkbook.Worksheets(1).range("D12:D13").ClearContents
    ThisWorkbook.Worksheets(1).range("G7:G9").ClearContents
    ThisWorkbook.Worksheets(1).range("G12:G13").ClearContents
    'PQ表の値クリア処理
    ThisWorkbook.Worksheets(1).range("D20:D21").ClearContents
    ThisWorkbook.Worksheets(1).range("D23:D24").ClearContents
    ThisWorkbook.Worksheets(1).range("G20:G21").ClearContents
    ThisWorkbook.Worksheets(1).range("G23:G24").ClearContents
       
    'シートの初期化
    Application.DisplayAlerts = False
    '”被保険者数集計”のワークシート以外を削除
    For Each Worksheet In Worksheets
        If Worksheet.Index <> 1 Then
            Worksheet.Delete
        End If
    Next
    Application.DisplayAlerts = True
    
    MsgBox "前回集計結果のリセットが完了しました。", vbInformation
End Sub
