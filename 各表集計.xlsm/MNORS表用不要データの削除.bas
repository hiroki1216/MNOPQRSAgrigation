Attribute VB_Name = "MNORS表用不要データの削除"
Sub MNORS表用不要データの削除()
Dim Worksheet As Worksheet 'ワークシート取得用
Dim i As Long '繰返し処理用
Dim lastcolumn As Long '繰返処理終点用
Dim standardRecord As Long 'レコード区分の列番号用
Dim standardGinushhi As Long '擬主区分の列番号用
Dim standardTokuteisya As Long '特定同一世帯所属者区分の列番号用
Dim standardKijyunnsousyotoku As Long '基準総所得（千円未満切捨）の列番号用

'各シートのカラム名を検索する処理
For Each Worksheet In Worksheets
    lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
    'カラムの検索
    For i = 1 To lastcolumn
        If Worksheet.Cells(1, i).Value = "レコード区分" Then
            standardRecord = i
            Worksheet.Cells(1, standardRecord).Interior.ColorIndex = 3
            'Debug.Print (standardRecord)
        ElseIf Worksheet.Cells(1, i).Value = "擬主区分" Then
            standardGinushhi = i
            Worksheet.Cells(1, standardGinushhi).Interior.ColorIndex = 3
            'Debug.Print (standardGinushhi)
        ElseIf Worksheet.Cells(1, i).Value = "特定同一世帯所属者区分" Then
            standardTokuteisya = i
            Worksheet.Cells(1, standardTokuteisya).Interior.ColorIndex = 3
            'Debug.Print (standardTokuteisya)
        ElseIf Worksheet.Cells(1, i).Value = "基準総所得（千円未満切捨）" Then
            standardKijyunnsousyotoku = i
            Worksheet.Cells(1, standardKijyunnsousyotoku).Interior.ColorIndex = 3
            'Debug.Print (standardTokuteisya)
        End If
    Next i
    
    '不要行の削除処理はここから
    
    Dim s As Long '繰返処理用
    Dim lastRow As Long '処理シートの最終行
    'A列の最終行を取得
    lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    'standardRecord=0の場合は、処理をスキップする
    If standardRecord = 0 Then
        GoTo Continue
    End If
    '不要行の削除処理
    'レコード区分
    For s = lastRow To 2 Step -1
        If Worksheet.Cells(s, standardRecord).Value = "世帯" Then
            Worksheet.Cells(s, standardRecord).EntireRow.Delete
            'Worksheet.Cells(s, standardRecord).Interior.ColorIndex = 3
        End If
    Next s
    
    '擬制世帯主区分
    lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    For s = lastRow To 2 Step -1
        If Worksheet.Cells(s, standardGinushhi).Value = "擬制世帯主" Then
            Worksheet.Cells(s, standardGinushhi).EntireRow.Delete
            'Worksheet.Cells(s, standardGinushhi).Interior.ColorIndex = 3
        End If
    Next s
    
    '特定同一世帯所属者区分
    lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    For s = lastRow To 2 Step -1
        If Worksheet.Cells(s, standardTokuteisya).Value = "特定同一世帯所属者" Then
            Worksheet.Cells(s, standardTokuteisya).EntireRow.Delete
            'Worksheet.Cells(s, standardTokuteisya).Interior.ColorIndex = 3
        End If
    Next s
    
    '基準総所得（千円未満切捨）
    lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    For s = lastRow To 2 Step -1
        If Worksheet.Cells(s, standardKijyunnsousyotoku).Value = 0 Then
            Worksheet.Cells(s, standardKijyunnsousyotoku).EntireRow.Delete
            'Worksheet.Cells(s, standardTokuteisya).Interior.ColorIndex = 3
        End If
    Next s
Continue:
Next
End Sub
