Attribute VB_Name = "ボツPQ表用M表かつO表対象者削除"
Sub PQ表用M表かつO表対象者削除()

Dim lastcolumn As Long 'カラム検索用最終列番号
Dim i As Long 'カラム検索繰返処理制御用
Dim standardAtenaMI As Long 'M表(医療)宛名番号カラム列番号用
Dim standardAtenaMK As Long 'M表(介護)宛名番号カラム列番号用
Dim standardAtenaN As Long 'N表宛名番号カラム列番号用
Dim standardAtenaO As Long 'O表宛名番号カラム列番号用
Dim standardAtenaP As Long 'P表宛名番号カラム列番号用
Dim standardAtenaQ As Long 'Q表宛名番号カラム列番号用

Dim standardSyubetuN As Long 'N表宛名番号カラム列番号用
Dim standardSyubetuO As Long 'O表宛名番号カラム列番号用
Dim standardSyubetuP As Long 'P表宛名番号カラム列番号用
Dim standardSyubetuQ As Long 'Q表宛名番号カラム列番号用

Dim WorksheetIndexMI As Long 'M表(医療)ワークシートインデックス用
Dim WorksheetIndexMK As Long 'M表(介護)ワークシートインデックス用
Dim WorksheetIndexN As Long 'N表ワークシートインデックス用
Dim WorksheetIndexO As Long 'O表ワークシートインデックス用
Dim WorksheetIndexP As Long 'P表ワークシートインデックス用
Dim WorksheetIndexQ As Long 'Q表ワークシートインデックス用

'ワークシートを再帰的に処理
For Each Worksheet In Worksheets
    '対象ワークシートの宛名番号カラムの列番号とワークシートのインデックス番号の取得
    If Worksheet.Name Like "*M表*" & "*医療*" Then
        WorksheetIndexMI = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "宛名番号" Then
                standardAtenaMI = i
                Worksheet.Cells(1, standardAtenaMI).Interior.ColorIndex = 3
            End If
        Next i
    ElseIf Worksheet.Name Like "*M表*" & "*介護*" Then
        WorksheetIndexMK = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "宛名番号" Then
                standardAtenaMK = i
                Worksheet.Cells(1, standardAtenaMK).Interior.ColorIndex = 3
            End If
        Next i
        
    ElseIf Worksheet.Name Like "*N表*" Then
        WorksheetIndexN = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "宛名番号" Then
                standardAtenaN = i
                Worksheet.Cells(1, standardAtenaN).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "保険税［料］種別" Then
                standardSyubetuN = i
                Worksheet.Cells(1, standardSyubetuN).Interior.ColorIndex = 3
            End If
        Next i
        
     ElseIf Worksheet.Name Like "*O表*" Then
        WorksheetIndexO = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "宛名番号" Then
                standardAtenaO = i
                Worksheet.Cells(1, standardAtenaO).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "保険税［料］種別" Then
                standardSyubetuO = i
                Worksheet.Cells(1, standardSyubetuO).Interior.ColorIndex = 3
            End If
        Next i
    
    ElseIf Worksheet.Name Like "*P表*" Then
        WorksheetIndexP = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "宛名番号" Then
                standardAtenaP = i
                Worksheet.Cells(1, standardAtenaP).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "保険税［料］種別" Then
                standardSyubetuP = i
                Worksheet.Cells(1, standardSyubetuP).Interior.ColorIndex = 3
            End If
        Next i
        
    ElseIf Worksheet.Name Like "*Q表*" Then
        WorksheetIndexQ = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "宛名番号" Then
                standardAtenaQ = i
                Worksheet.Cells(1, standardAtenaQ).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "保険税［料］種別" Then
                standardSyubetuQ = i
                Worksheet.Cells(1, standardSyubetuQ).Interior.ColorIndex = 3
            End If
        Next i
    Else
        GoTo Continue
    End If
Continue:
Next

'Debug.Print WorksheetIndexMI
'Debug.Print WorksheetIndexMK
'Debug.Print WorksheetIndexN
'Debug.Print WorksheetIndexO
'Debug.Print WorksheetIndexP
'Debug.Print WorksheetIndexQ
'
'Debug.Print standardAtenaMI
'Debug.Print standardAtenaMK
'Debug.Print standardAtenaN
'Debug.Print standardAtenaO
'Debug.Print standardAtenaP
'Debug.Print standardAtenaQ


'M表かつO表対象者を削除はここから
Dim m As Long 'M表繰返処理制御用
Dim o As Long 'O表繰返処理制御用
Dim lastRowMI As Long 'M表(医療)繰返処理制御用
Dim lastRowMK As Long 'M表(介護)繰返処理制御用
Dim lastRowO As Long 'O表繰返処理制御用


lastRowMI = ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(1, 1).End(xlDown).Row

lastRowMK = ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(1, 1).End(xlDown).Row

lastRowO = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(1, 1).End(xlDown).Row



'M表(医療分)かつO表処理
For o = lastRowO To 2 Step -1
    For m = lastRowMI To 2 Step -1
        If ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).Value = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardAtenaO).Value And ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardSyubetuO).Value = "医療分" Then
            'ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).EntireRow.Delete
            ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).Interior.ColorIndex = 3
        End If
    Next m
Next o
'M表(介護分)かつO表処理
For o = lastRowO To 2 Step -1
    For m = lastRowMK To 2 Step -1
        If ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(m, standardAtenaMK).Value = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardAtenaO).Value And ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardSyubetuO).Value = "介護分" Then
            'ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(m, standardAtenaMK).EntireRow.Delete
            ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(m, standardAtenaMK).Interior.ColorIndex = 3
        End If
    Next m
Next o

'PQ表集計はここから
'Dim counterP As Long 'P表被保険者数カウンター
'Dim counterQ As Long 'Q表被保険者数カウンター
'Dim pq As Long 'PQ表繰返処理制御用
'Dim n As Long 'N繰返処理制御用
'Dim lastRowP As Long 'P表繰返処理制御用
'Dim lastRowQ As Long 'Q表繰返処理制御用
'Dim lastRowN As Long 'N表繰返処理制御用
'
'lastRowP = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(1, 1).End(xlDown).Row
'lastRowQ = ThisWorkbook.Worksheets(WorksheetIndexQ).Cells(1, 1).End(xlDown).Row
'lastRowN = ThisWorkbook.Worksheets(WorksheetIndexN).Cells(1, 1).End(xlDown).Row
'lastRowMI = ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(1, 1).End(xlDown).Row
'lastRowMK = ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(1, 1).End(xlDown).Row

'P表集計(医療分)
'counterP = 0
'For m = 2 To lastRowMI
'    For pq = 2 To lastRowP
'        If ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).Value = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(pq, standardAtenaP).Value And ThisWorkbook.Worksheets(WorksheetIndexP).Cells(pq, standardSyubetuP).Value = "医療分" And ThisWorkbook.Worksheets(WorksheetIndexP).Cells(pq, standardSyubetuP).Value > 0 Then
'            counterP = counterP + 1
'        End If
'    Next pq
'Next m
'Debug.Print counterP
End Sub
