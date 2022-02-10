Attribute VB_Name = "PQ表集計配列使用版"
Sub PQ表集計配列使用版()
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
Dim myArray As Variant '配列用

'ワークシートを再帰的に処理
For Each Worksheet In Worksheets

    Dim range As Variant '配列格納範囲用
    
    '対象ワークシートの宛名番号カラムの列番号とワークシートのインデックス番号の取得
    If Worksheet.Name Like "*M表*" & "*医療*" Then
        Worksheet.Select
        WorksheetIndexMI = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '検索用２次元配列の定義
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "宛名番号" Then
                standardAtenaMI = i
                Worksheet.Cells(1, standardAtenaMI).Interior.ColorIndex = 3
            End If
        Next i
        '配列の初期化
        Erase myArray
        
    ElseIf Worksheet.Name Like "*M表*" & "*介護*" Then
        Worksheet.Select
        WorksheetIndexMK = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '検索用２次元配列の定義
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "宛名番号" Then
                standardAtenaMK = i
                Worksheet.Cells(1, standardAtenaMK).Interior.ColorIndex = 3
            End If
        Next i
        '配列の初期化
        Erase myArray

    ElseIf Worksheet.Name Like "*N表*" Then
        Worksheet.Select
        WorksheetIndexN = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '検索用２次元配列の定義
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "宛名番号" Then
                standardAtenaN = i
                Worksheet.Cells(1, standardAtenaN).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "保険税［料］種別" Then
                standardSyubetuN = i
                Worksheet.Cells(1, standardSyubetuN).Interior.ColorIndex = 3
            End If
        Next i
        '配列の初期化
        Erase myArray
        
     ElseIf Worksheet.Name Like "*O表*" Then
        Worksheet.Select
        WorksheetIndexO = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '検索用２次元配列の定義
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "宛名番号" Then
                standardAtenaO = i
                Worksheet.Cells(1, standardAtenaO).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "保険税［料］種別" Then
                standardSyubetuO = i
                Worksheet.Cells(1, standardSyubetuO).Interior.ColorIndex = 3
            End If
        Next i
        '配列の初期化
        Erase myArray
    
    ElseIf Worksheet.Name Like "*P表*" Then
        Worksheet.Select
        WorksheetIndexP = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '検索用２次元配列の定義
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "宛名番号" Then
                standardAtenaP = i
                Worksheet.Cells(1, standardAtenaP).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "保険税［料］種別" Then
                standardSyubetuP = i
                Worksheet.Cells(1, standardSyubetuP).Interior.ColorIndex = 3
            End If
        Next i
        '配列の初期化
        Erase myArray
        
    ElseIf Worksheet.Name Like "*Q表*" Then
        Worksheet.Select
        WorksheetIndexQ = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '検索用２次元配列の定義
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "宛名番号" Then
                standardAtenaQ = i
                Worksheet.Cells(1, standardAtenaQ).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "保険税［料］種別" Then
                standardSyubetuQ = i
                Worksheet.Cells(1, standardSyubetuQ).Interior.ColorIndex = 3
            End If
        Next i
        '配列の初期化
        Erase myArray
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
'
'Debug.Print standardSyubetuO
'Debug.Print standardSyubetuP
'Debug.Print standardSyubetuQ



'(M-O)表データ抽出処理はここから----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim m As Long 'M表繰返処理制御用
Dim o As Long 'O表繰返処理制御用

'M表(医療分)かつO表処理
Dim myArray2 As Variant 'M表(医療)データ格納用配列（宛名列）
Dim myArray3 As Variant  'O表データ格納用配列（宛名列）
Dim myArray4 As Variant  'O表データ格納用配列（保険税［料］種別）
Dim myArray5 As Variant 'M表(介護)データ格納用配列（宛名列）

Dim lastRow2 As Long 'M表(医療)データ配列格納範囲（宛名列）最終行番号
Dim lastRow3 As Long  'O表データ配列格納範囲（宛名列）最終行番号
Dim lastRow4 As Long  'O表データ配列格納範囲（保険税［料］種別）最終行番号
Dim lastRow5 As Long 'M表(介護)データ配列格納範囲（宛名列）最終行番号

Dim range2 As range 'M表(医療)データ配列格納範囲（宛名列）
Dim range3 As range 'O表データ配列格納範囲（宛名列）
Dim range4 As range 'O表データ配列格納範囲（保険税［料］種別）
Dim range5 As range 'M表(介護)データ配列格納範囲（宛名列）

'範囲最終行番号の取得/配列格納データ範囲の取得

'M表医療（宛名）
ThisWorkbook.Worksheets(WorksheetIndexMI).Select
lastRow2 = ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(1, 1).End(xlDown).Row
Set range2 = ThisWorkbook.Worksheets(WorksheetIndexMI).range(Cells(2, standardAtenaMI), Cells(lastRow2, standardAtenaMI))

'M表介護（宛名）
ThisWorkbook.Worksheets(WorksheetIndexMK).Select
lastRow5 = ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(1, 1).End(xlDown).Row
Set range5 = ThisWorkbook.Worksheets(WorksheetIndexMK).range(Cells(2, standardAtenaMK), Cells(lastRow5, standardAtenaMK))

'O表（宛名）
ThisWorkbook.Worksheets(WorksheetIndexO).Select
lastRow3 = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(1, 1).End(xlDown).Row
Set range3 = ThisWorkbook.Worksheets(WorksheetIndexO).range(Cells(2, standardAtenaO), Cells(lastRow3, standardAtenaO))

'O表（保険税［料］種別）
ThisWorkbook.Worksheets(WorksheetIndexO).Select
lastRow4 = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(1, 1).End(xlDown).Row
Set range4 = ThisWorkbook.Worksheets(WorksheetIndexO).range(Cells(2, standardSyubetuO), Cells(lastRow4, standardSyubetuO))


'各配列(２次元配列)にデータを格納
myArray2 = (range2) 'M表(医療)データ配列
myArray3 = (range3) 'O表(宛名)データ配列
myArray4 = (range4) 'O表(種別)データ配列
myArray5 = (range5) 'M表(介護)データ配列


'M表（医療）からO表(医療)を除く配列(M-O)を作成。（要素数は削除できないため、判定済みの前の要素で上書き）
Dim testArray As Variant
ReDim testArray(1 To 10000, 1 To 1)
Dim t As Long '配列の列番号用
t = 1 '列番号の初期化
Dim counterTest As Long '重複カウンター
counterTest = 0 '重複カウンターの初期化

For m = LBound(myArray2, 1) To UBound(myArray2, 1)
    For o = LBound(myArray3, 1) To UBound(myArray3, 1)
        If myArray2(m, 1) = myArray3(o, 1) And myArray4(o, 1) = "医療分" Then
            'M表（医療）かつO表のデータを配列に格納
            testArray(t, 1) = myArray2(m, 1)
            '判定済の要素で上書き
             myArray2(m, 1) = myArray2(m - 1, 1)
            counterTest = counterTest + 1
            'If分の結果がTRUEであった場合、カウンターに１を加え、testArrayにインクリメンタルに値を格納していく
            t = t + 1
        End If
    Next o
Next m
Debug.Print UBound(myArray2)

'重複削除とM表かつO表データ確認のため、シートを新規作成”
Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "M表かつO表"
'配列の重複を削除するために一度シートに出力
Worksheets("M表かつO表").range("B1:B" & UBound(myArray2, 1)) = myArray2
'シートに出力された値の重複を削除
Worksheets("M表かつO表").range("B1:B" & UBound(myArray2, 1)).RemoveDuplicates Columns:=1, Header:=xlNo
'重複を削除したデータを配列に再格納
Erase myArray2
Worksheets("M表かつO表").Select
lastRow2 = Worksheets("M表かつO表").Cells(1, 2).End(xlDown).Row '最終行番号の取得
Set range2 = Worksheets("M表かつO表").range(Cells(1, 2), Cells(lastRow2, 2)) '配列に格納する値のセル範囲を選択
myArray2 = (range2) '配列に格納
'M表(医療)かつO表(医療)の対象者を出力
Worksheets("M表かつO表").range("A1:A" & UBound(testArray)) = testArray
Debug.Print UBound(myArray2)
Debug.Print counterTest

'M表(介護)処理用に変数を初期化
t = 1 '列番号の初期化
counterTest = 0 '重複カウンターの初期化
Erase testArray '配列の初期化

ReDim testArray(1 To 10000, 1 To 1)

'M表（介護）からO表(介護)を除く配列(M-O)を作成。（要素数は削除できないため、判定済みの前の要素で上書き）

For m = LBound(myArray5, 1) To UBound(myArray5, 1)
    For o = LBound(myArray3, 1) To UBound(myArray3, 1)
        If myArray5(m, 1) = myArray3(o, 1) And myArray4(o, 1) = "介護分" Then
            'M表（医療）かつO表のデータを配列に格納
            testArray(t, 1) = myArray5(m, 1)
            '判定済の要素で上書き
             myArray5(m, 1) = myArray5(m - 1, 1)
            counterTest = counterTest + 1
            'If分の結果がTRUEであった場合、カウンターに１を加え、testArrayにインクリメンタルに値を格納していく
            t = t + 1
        End If
    Next o
Next m
Debug.Print UBound(myArray5)

'配列の重複を削除するために一度シートに出力
Worksheets("M表かつO表").range("D1:D" & UBound(myArray5, 1)) = myArray5
'シートに出力された値の重複を削除
Worksheets("M表かつO表").range("D1:D" & UBound(myArray5, 1)).RemoveDuplicates Columns:=1, Header:=xlNo
'重複を削除したデータを配列に再格納
Erase myArray5 '配列の初期化
Worksheets("M表かつO表").Select
lastRow5 = Worksheets("M表かつO表").Cells(1, 4).End(xlDown).Row '最終行番号の取得
Set range5 = Worksheets("M表かつO表").range(Cells(1, 4), Cells(lastRow5, 4)) '配列に格納する値のセル範囲を選択
myArray5 = (range5) '配列に格納
'M表(介護)かつO表(介護)の対象者を出力
Worksheets("M表かつO表").range("C1:C" & UBound(testArray)) = testArray
Debug.Print UBound(myArray5)
Debug.Print counterTest
'testArrayの初期化
Erase testArray
'O表関連の配列を削除
Erase myArray3
Erase myArray4

'M表かつPQ表集計処理はここから----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim myArray6 As Variant 'P表データ格納用配列（宛名列）
Dim myArray7 As Variant 'P表データ格納用配列（保険税［料］種別）
Dim myArray8 As Variant 'Q表データ格納用配列（宛名列）
Dim myArray9 As Variant 'Q表データ格納用配列（保険税［料］種別）

Dim lastRow6 As Long 'P表データ格納用配列（宛名列）最終行番号
Dim lastRow7 As Long 'P表データ格納用配列（保険税［料］種別）最終行番号
Dim lastRow8 As Long 'Q表データ格納用配列（宛名列）最終行番号
Dim lastRow9 As Long 'Q表データ格納用配列（保険税［料］種別）最終行番号

Dim range6 As range 'P表データ配列格納範囲（宛名列）
Dim range7 As range 'P表データ配列格納範囲（保険税［料］種別）
Dim range8 As range 'Q表データ配列格納範囲（宛名列）
Dim range9 As range 'Q表データ配列格納範囲（保険税［料］種別）

'範囲最終行番号の取得/配列格納データ範囲の取得

'P表（宛名）
ThisWorkbook.Worksheets(WorksheetIndexP).Select
lastRow6 = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(1, 1).End(xlDown).Row
Set range6 = ThisWorkbook.Worksheets(WorksheetIndexP).range(Cells(2, standardAtenaP), Cells(lastRow6, standardAtenaP))

'P表（保険税［料］種別）
lastRow7 = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(1, 1).End(xlDown).Row
Set range7 = ThisWorkbook.Worksheets(WorksheetIndexP).range(Cells(2, standardSyubetuP), Cells(lastRow7, standardSyubetuP))

'O表（宛名）
ThisWorkbook.Worksheets(WorksheetIndexQ).Select
lastRow8 = ThisWorkbook.Worksheets(WorksheetIndexQ).Cells(1, 1).End(xlDown).Row
Set range8 = ThisWorkbook.Worksheets(WorksheetIndexQ).range(Cells(2, standardAtenaQ), Cells(lastRow8, standardAtenaQ))

'O表（保険税［料］種別）
lastRow9 = ThisWorkbook.Worksheets(WorksheetIndexQ).Cells(1, 1).End(xlDown).Row
Set range9 = ThisWorkbook.Worksheets(WorksheetIndexQ).range(Cells(2, standardSyubetuQ), Cells(lastRow9, standardSyubetuQ))


'各配列(２次元配列)にデータを格納
myArray6 = (range6) 'P表(宛名)データ配列
myArray7 = (range7) 'P表(種別)データ配列
myArray8 = (range8) 'Q表(宛名)データ配列
myArray9 = (range9) 'Q表(種別)データ配列

'繰返処理制御用
Dim pq As Long 'PQ表用

'(M-O)表(医療)かつP表(医療)集計用カウンター
Dim counterMPI As Long
counterMPI = 0

'(M-O)表(介護)かつP表(介護)集計用カウンター
Dim counterMPK As Long
counterMPK = 0

'(M-O)表(医療)かつQ表(医療)集計用カウンター
Dim counterMQI As Long
counterMQI = 0

'(M-O)表(介護)かつQ表(介護)集計用カウンター
Dim counterMQK As Long
counterMQK = 0


'(M-O)表(医療)かつP表(医療)集計
For m = LBound(myArray2, 1) To UBound(myArray2, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray2(m, 1) = myArray6(pq, 1) And myArray7(pq, 1) = "医療分" Then
            counterMPI = counterMPI + 1
        End If
    Next pq
Next m
Debug.Print counterMPI


'(M-O)表(介護)かつP表(介護)集計
For m = LBound(myArray5, 1) To UBound(myArray5, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray5(m, 1) = myArray6(pq, 1) And myArray7(pq, 1) = "介護分" Then
            counterMPK = counterMPK + 1
        End If
    Next pq
Next m
Debug.Print counterMPK


'(M-O)表(医療)かつQ表(医療)集計
For m = LBound(myArray2, 1) To UBound(myArray2, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray2(m, 1) = myArray8(pq, 1) And myArray9(pq, 1) = "医療分" Then
            counterMQI = counterMQI + 1
        End If
    Next pq
Next m
Debug.Print counterMQI


'(M-O)表(介護)かつQ表(介護)集計
For m = LBound(myArray5, 1) To UBound(myArray5, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray5(m, 1) = myArray8(pq, 1) And myArray9(pq, 1) = "介護分" Then
            counterMQK = counterMQK + 1
        End If
    Next pq
Next m
Debug.Print counterMQK


'M表関連の配列を初期化
Erase myArray2 'PQ表の集計が終わるまで残す
Erase myArray5 'PQ表の集計が終わるまで残す

'N表かつPQ表集計処理はここから----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim myArray10 As Variant 'N表データ格納用配列（宛名列）
Dim myArray11 As Variant 'N表データ格納用配列（保険税［料］種別）

Dim lastRow10 As Long 'N表データ格納用配列（宛名列）最終行番号
Dim lastRow11 As Long 'N表データ格納用配列（保険税［料］種別）最終行番号

Dim range10 As range 'N表データ配列格納範囲（宛名列）
Dim range11 As range 'N表データ配列格納範囲（保険税［料］種別）

'繰返処理制御用
Dim n As Long 'N表用

'N表(医療)かつP表(医療)集計用カウンター
Dim counterNPI As Long
counterNPI = 0

'N表(介護)かつP表(介護)集計用カウンター
Dim counterNPK As Long
counterNPK = 0

'N表(医療)かつQ表(医療)集計用カウンター
Dim counterNQI As Long
counterNQI = 0

'N表(介護)かつQ表(介護)集計用カウンター
Dim counterNQK As Long
counterNQK = 0


'範囲最終行番号の取得/配列格納データ範囲の取得

'N表（宛名）
ThisWorkbook.Worksheets(WorksheetIndexN).Select
lastRow10 = ThisWorkbook.Worksheets(WorksheetIndexN).Cells(1, 1).End(xlDown).Row
Set range10 = ThisWorkbook.Worksheets(WorksheetIndexN).range(Cells(2, standardAtenaN), Cells(lastRow10, standardAtenaN))

'N表（保険税［料］種別）
lastRow11 = ThisWorkbook.Worksheets(WorksheetIndexN).Cells(1, 1).End(xlDown).Row
Set range11 = ThisWorkbook.Worksheets(WorksheetIndexN).range(Cells(2, standardSyubetuN), Cells(lastRow11, standardSyubetuN))

'各配列(２次元配列)にデータを格納
myArray10 = (range10) 'N表(宛名)データ配列
myArray11 = (range11) 'N表(種別)データ配列

'N表(医療)かつP表(医療)集計
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray10(n, 1) = myArray6(pq, 1) And myArray11(n, 1) = "医療分" And myArray7(pq, 1) = "医療分" Then
            counterNPI = counterNPI + 1
        End If
    Next pq
Next n
Debug.Print counterNPI

'N表(介護)かつP表(介護)集計
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray10(n, 1) = myArray6(pq, 1) And myArray11(n, 1) = "介護分" And myArray7(pq, 1) = "介護分" Then
            counterNPK = counterNPK + 1
        End If
    Next pq
Next n
Debug.Print counterNPK

'N表(医療)かつQ表(医療)集計
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray10(n, 1) = myArray8(pq, 1) And myArray11(n, 1) = "医療分" And myArray9(pq, 1) = "医療分" Then
            counterNQI = counterNQI + 1
        End If
    Next pq
Next n
Debug.Print counterNQI

'N表(介護)かつQ表(介護)集計
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray10(n, 1) = myArray8(pq, 1) And myArray11(n, 1) = "介護分" And myArray9(pq, 1) = "介護分" Then
            counterNQK = counterNQK + 1
        End If
    Next pq
Next n
Debug.Print counterNQK

'PQN表の配列を初期化
Erase myArray6
Erase myArray7
Erase myArray8
Erase myArray9
Erase myArray10
Erase myArray11

'結果の出力はここから----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

ThisWorkbook.Worksheets(1).Select
ThisWorkbook.Worksheets(1).range("D20").Value = counterMPI '(M-O)表(医療)かつP表(医療)集計結果を出力
ThisWorkbook.Worksheets(1).range("G20").Value = counterMPK '(M-O)表(介護)かつP表(介護)集計結果を出力
ThisWorkbook.Worksheets(1).range("D23").Value = counterMQI '(M-O)表(医療)かつQ表(医療)集計結果を出力
ThisWorkbook.Worksheets(1).range("G23").Value = counterMQK '(M-O)表(介護)かつQ表(介護)集計結果を出力
ThisWorkbook.Worksheets(1).range("D21").Value = counterNPI 'N表(医療)かつP表(医療)集計結果を出力
ThisWorkbook.Worksheets(1).range("G21").Value = counterNPK 'N表(介護)かつP表(介護)集計結果を出力
ThisWorkbook.Worksheets(1).range("D24").Value = counterNQI 'N表(医療)かつP表(医療)集計結果を出力
ThisWorkbook.Worksheets(1).range("G24").Value = counterNQK 'N表(介護)かつP表(介護)集計結果を出力
End Sub
