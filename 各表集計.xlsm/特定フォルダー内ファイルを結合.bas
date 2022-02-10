Attribute VB_Name = "特定フォルダー内ファイルを結合"
Sub 特定フォルダー内ファイルを結合()
On Error Resume Next

'集計ボタンを複数回押されたときの処理（シートの初期化）
Application.DisplayAlerts = False
Worksheets("merge").Delete
Application.DisplayAlerts = True

'”mergeファイルを作成”
Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "merge"

Dim folderPath
folderPath = ThisWorkbook.Worksheets(1).range("B2").Value '処理フォルダーを指定

'拡張子の種類で場合分け
Dim fileType
If ThisWorkbook.Worksheets(1).range("B1").Value = "Excel" Then
    fileType = "\*.xlsx*"
Else
    fileType = "\*.csv"
End If

Dim mergeWorkbook
'統合するファイル(excelでいうところのbook)を取得
mergeWorkbook = Dir(folderPath & fileType) '処理ファイル名を取得(Dir関数は、戻り値としてファイル名を返す)

'最初の行をコピぺ
 Workbooks.Open Filename:=folderPath & "\" & mergeWorkbook '処理フォルダー\処理ファイル名を開く
 Workbooks(mergeWorkbook).Worksheets(1).Rows(1).Copy ThisWorkbook.Worksheets("merge").range("A1")
 Application.DisplayAlerts = False
 Workbooks(mergeWorkbook).Close
 Application.DisplayAlerts = True
 
 'データ識別用のファイル名カラムを最終列に追加
 Dim lastcolumn As Long
 lastcolumn = ThisWorkbook.Worksheets("merge").Cells(1, Columns.Count).End(xlToLeft).Column
 ThisWorkbook.Worksheets("merge").Cells(1, lastcolumn + 1).Value = "ファイル名"
 
 Dim counter As Long
 counter = 0
 

'検索結果が無くなるまで繰返処理
Do Until mergeWorkbook = ""
    Workbooks.Open Filename:=folderPath & "\" & mergeWorkbook '処理フォルダー\処理ファイル名を開く
    If Workbooks.Count = 20 Then Exit Do

    Dim mergeWorkbookdata
    Dim thisWorkbookdata
    Dim i

    '開いたエクセルファイルに複数のシートが存在する場合は、シート分、繰返処理を行う
    For i = 1 To Workbooks(mergeWorkbook).Worksheets.Count

    'マージされるワークブックのA列の最終行
    mergeWorkbookdata = Workbooks(mergeWorkbook).Worksheets(i).range("a" & Rows.Count).End(xlUp).Row
    'マージするワークブックのA列の最終行
    thisWorkbookdata = ThisWorkbook.Worksheets("merge").range("a" & Rows.Count).End(xlUp).Row
    
    'マージするファイルの最終行のに続けて、マージされるワークブックをコピペ
    Workbooks(mergeWorkbook).Worksheets(i).Rows("2:" & mergeWorkbookdata).Copy ThisWorkbook.Worksheets("merge").range("a" & thisWorkbookdata + 1)
    
    
    Next i
    
     'ファイル名を追加
    Dim lastRow As Long
    Dim startrow As Long
    Dim s As Long
    
    'A列最終行番号
    lastRow = ThisWorkbook.Worksheets("merge").Cells(1, 1).End(xlDown).Row
    '最終列番号
    lastcolumn = ThisWorkbook.Worksheets("merge").Cells(1, Columns.Count).End(xlToLeft).Column
    '最終列入力済み最終行番号
    
    'ファイル名出力処理を統合回数で場合分け（１回目のみ出力がないため『２』が入る）
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
    
    'ファイルが統合された回数分加算する
    counter = counter + 1
Loop



End Sub

