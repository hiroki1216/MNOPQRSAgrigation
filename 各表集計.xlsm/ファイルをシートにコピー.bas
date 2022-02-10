Attribute VB_Name = "ファイルをシートにコピー"
Sub ファイルをシートにコピー()

Dim Worksheet As Worksheet 'ワークシート取得用（初期化処理）
Dim folderPath As String '処理フォルダーのディレクトリを取得用
Dim fileType As String 'ファイルの拡張子用
Dim mergeWorkbook As String '処理ファイル用
Dim mergeWorkbookdata As Long '処理ファイルのA列最終行取得用

'集計ボタンを複数回押されたときの処理（シートの初期化）
Application.DisplayAlerts = False
'”被保険者数集計”のワークシート以外を削除
For Each Worksheet In Worksheets
    If Worksheet.Index <> 1 Then
        Worksheet.Delete
    End If
Next
Application.DisplayAlerts = True

'処理フォルダーを指定
folderPath = ThisWorkbook.Worksheets(1).range("B2").Value

'セルの値で拡張子を場合分け
If Worksheets(1).range("B1").Value = "Excel" Then
    fileType = "\*.xlsx*"
Else
    fileType = "\*.csv"
End If

'処理対象ファイル名の取得
mergeWorkbook = Dir(folderPath & fileType) '(Dir関数は、戻り値に、文字列型のファイル名を返す)

'対象ファイルを統合用のファイルのシートにコピー
Do Until mergeWorkbook = ""
    
    'マージするワークブックにマージされるファイル名のシートを新規作成”
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = mergeWorkbook
    '処理フォルダー\処理ファイル名を開く
    Workbooks.Open Filename:=folderPath & "\" & mergeWorkbook
    
    'マージされるワークブックのA列の最終行を取得
    mergeWorkbookdata = Workbooks(mergeWorkbook).Worksheets(1).Cells(1, 1).End(xlDown).Row
    
    'マージするファイルの最終行にマージされるワークブックをコピペ
    Workbooks(mergeWorkbook).Worksheets(1).Rows("1:" & mergeWorkbookdata).Copy ThisWorkbook.Worksheets(mergeWorkbook).range("A1")
    Application.DisplayAlerts = False
    Workbooks(mergeWorkbook).Close
    Application.DisplayAlerts = True
    
    '指定フォルダー内の対象ファイルを再帰的に処理
    mergeWorkbook = Dir()

Loop

End Sub
