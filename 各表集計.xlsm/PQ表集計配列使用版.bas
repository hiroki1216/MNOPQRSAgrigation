Attribute VB_Name = "PQ�\�W�v�z��g�p��"
Sub PQ�\�W�v�z��g�p��()
Dim lastcolumn As Long '�J���������p�ŏI��ԍ�
Dim i As Long '�J���������J�ԏ�������p
Dim standardAtenaMI As Long 'M�\(���)�����ԍ��J������ԍ��p
Dim standardAtenaMK As Long 'M�\(���)�����ԍ��J������ԍ��p
Dim standardAtenaN As Long 'N�\�����ԍ��J������ԍ��p
Dim standardAtenaO As Long 'O�\�����ԍ��J������ԍ��p
Dim standardAtenaP As Long 'P�\�����ԍ��J������ԍ��p
Dim standardAtenaQ As Long 'Q�\�����ԍ��J������ԍ��p

Dim standardSyubetuN As Long 'N�\�����ԍ��J������ԍ��p
Dim standardSyubetuO As Long 'O�\�����ԍ��J������ԍ��p
Dim standardSyubetuP As Long 'P�\�����ԍ��J������ԍ��p
Dim standardSyubetuQ As Long 'Q�\�����ԍ��J������ԍ��p

Dim WorksheetIndexMI As Long 'M�\(���)���[�N�V�[�g�C���f�b�N�X�p
Dim WorksheetIndexMK As Long 'M�\(���)���[�N�V�[�g�C���f�b�N�X�p
Dim WorksheetIndexN As Long 'N�\���[�N�V�[�g�C���f�b�N�X�p
Dim WorksheetIndexO As Long 'O�\���[�N�V�[�g�C���f�b�N�X�p
Dim WorksheetIndexP As Long 'P�\���[�N�V�[�g�C���f�b�N�X�p
Dim WorksheetIndexQ As Long 'Q�\���[�N�V�[�g�C���f�b�N�X�p
Dim myArray As Variant '�z��p

'���[�N�V�[�g���ċA�I�ɏ���
For Each Worksheet In Worksheets

    Dim range As Variant '�z��i�[�͈͗p
    
    '�Ώۃ��[�N�V�[�g�̈����ԍ��J�����̗�ԍ��ƃ��[�N�V�[�g�̃C���f�b�N�X�ԍ��̎擾
    If Worksheet.Name Like "*M�\*" & "*���*" Then
        Worksheet.Select
        WorksheetIndexMI = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '�����p�Q�����z��̒�`
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�����ԍ�" Then
                standardAtenaMI = i
                Worksheet.Cells(1, standardAtenaMI).Interior.ColorIndex = 3
            End If
        Next i
        '�z��̏�����
        Erase myArray
        
    ElseIf Worksheet.Name Like "*M�\*" & "*���*" Then
        Worksheet.Select
        WorksheetIndexMK = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '�����p�Q�����z��̒�`
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�����ԍ�" Then
                standardAtenaMK = i
                Worksheet.Cells(1, standardAtenaMK).Interior.ColorIndex = 3
            End If
        Next i
        '�z��̏�����
        Erase myArray

    ElseIf Worksheet.Name Like "*N�\*" Then
        Worksheet.Select
        WorksheetIndexN = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '�����p�Q�����z��̒�`
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�����ԍ�" Then
                standardAtenaN = i
                Worksheet.Cells(1, standardAtenaN).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�ی��Łm���n���" Then
                standardSyubetuN = i
                Worksheet.Cells(1, standardSyubetuN).Interior.ColorIndex = 3
            End If
        Next i
        '�z��̏�����
        Erase myArray
        
     ElseIf Worksheet.Name Like "*O�\*" Then
        Worksheet.Select
        WorksheetIndexO = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '�����p�Q�����z��̒�`
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�����ԍ�" Then
                standardAtenaO = i
                Worksheet.Cells(1, standardAtenaO).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�ی��Łm���n���" Then
                standardSyubetuO = i
                Worksheet.Cells(1, standardSyubetuO).Interior.ColorIndex = 3
            End If
        Next i
        '�z��̏�����
        Erase myArray
    
    ElseIf Worksheet.Name Like "*P�\*" Then
        Worksheet.Select
        WorksheetIndexP = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '�����p�Q�����z��̒�`
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�����ԍ�" Then
                standardAtenaP = i
                Worksheet.Cells(1, standardAtenaP).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�ی��Łm���n���" Then
                standardSyubetuP = i
                Worksheet.Cells(1, standardSyubetuP).Interior.ColorIndex = 3
            End If
        Next i
        '�z��̏�����
        Erase myArray
        
    ElseIf Worksheet.Name Like "*Q�\*" Then
        Worksheet.Select
        WorksheetIndexQ = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        Set range = Worksheet.range(Cells(1, 1), Cells(1, lastcolumn))
        '�����p�Q�����z��̒�`
        myArray = (range)
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�����ԍ�" Then
                standardAtenaQ = i
                Worksheet.Cells(1, standardAtenaQ).Interior.ColorIndex = 3
            End If
        Next i
        
        
        For i = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(1, i) = "�ی��Łm���n���" Then
                standardSyubetuQ = i
                Worksheet.Cells(1, standardSyubetuQ).Interior.ColorIndex = 3
            End If
        Next i
        '�z��̏�����
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



'(M-O)�\�f�[�^���o�����͂�������----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim m As Long 'M�\�J�ԏ�������p
Dim o As Long 'O�\�J�ԏ�������p

'M�\(��Õ�)����O�\����
Dim myArray2 As Variant 'M�\(���)�f�[�^�i�[�p�z��i������j
Dim myArray3 As Variant  'O�\�f�[�^�i�[�p�z��i������j
Dim myArray4 As Variant  'O�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj
Dim myArray5 As Variant 'M�\(���)�f�[�^�i�[�p�z��i������j

Dim lastRow2 As Long 'M�\(���)�f�[�^�z��i�[�͈́i������j�ŏI�s�ԍ�
Dim lastRow3 As Long  'O�\�f�[�^�z��i�[�͈́i������j�ŏI�s�ԍ�
Dim lastRow4 As Long  'O�\�f�[�^�z��i�[�͈́i�ی��Łm���n��ʁj�ŏI�s�ԍ�
Dim lastRow5 As Long 'M�\(���)�f�[�^�z��i�[�͈́i������j�ŏI�s�ԍ�

Dim range2 As range 'M�\(���)�f�[�^�z��i�[�͈́i������j
Dim range3 As range 'O�\�f�[�^�z��i�[�͈́i������j
Dim range4 As range 'O�\�f�[�^�z��i�[�͈́i�ی��Łm���n��ʁj
Dim range5 As range 'M�\(���)�f�[�^�z��i�[�͈́i������j

'�͈͍ŏI�s�ԍ��̎擾/�z��i�[�f�[�^�͈͂̎擾

'M�\��Ái�����j
ThisWorkbook.Worksheets(WorksheetIndexMI).Select
lastRow2 = ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(1, 1).End(xlDown).Row
Set range2 = ThisWorkbook.Worksheets(WorksheetIndexMI).range(Cells(2, standardAtenaMI), Cells(lastRow2, standardAtenaMI))

'M�\���i�����j
ThisWorkbook.Worksheets(WorksheetIndexMK).Select
lastRow5 = ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(1, 1).End(xlDown).Row
Set range5 = ThisWorkbook.Worksheets(WorksheetIndexMK).range(Cells(2, standardAtenaMK), Cells(lastRow5, standardAtenaMK))

'O�\�i�����j
ThisWorkbook.Worksheets(WorksheetIndexO).Select
lastRow3 = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(1, 1).End(xlDown).Row
Set range3 = ThisWorkbook.Worksheets(WorksheetIndexO).range(Cells(2, standardAtenaO), Cells(lastRow3, standardAtenaO))

'O�\�i�ی��Łm���n��ʁj
ThisWorkbook.Worksheets(WorksheetIndexO).Select
lastRow4 = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(1, 1).End(xlDown).Row
Set range4 = ThisWorkbook.Worksheets(WorksheetIndexO).range(Cells(2, standardSyubetuO), Cells(lastRow4, standardSyubetuO))


'�e�z��(�Q�����z��)�Ƀf�[�^���i�[
myArray2 = (range2) 'M�\(���)�f�[�^�z��
myArray3 = (range3) 'O�\(����)�f�[�^�z��
myArray4 = (range4) 'O�\(���)�f�[�^�z��
myArray5 = (range5) 'M�\(���)�f�[�^�z��


'M�\�i��Áj����O�\(���)�������z��(M-O)���쐬�B�i�v�f���͍폜�ł��Ȃ����߁A����ς݂̑O�̗v�f�ŏ㏑���j
Dim testArray As Variant
ReDim testArray(1 To 10000, 1 To 1)
Dim t As Long '�z��̗�ԍ��p
t = 1 '��ԍ��̏�����
Dim counterTest As Long '�d���J�E���^�[
counterTest = 0 '�d���J�E���^�[�̏�����

For m = LBound(myArray2, 1) To UBound(myArray2, 1)
    For o = LBound(myArray3, 1) To UBound(myArray3, 1)
        If myArray2(m, 1) = myArray3(o, 1) And myArray4(o, 1) = "��Õ�" Then
            'M�\�i��Áj����O�\�̃f�[�^��z��Ɋi�[
            testArray(t, 1) = myArray2(m, 1)
            '����ς̗v�f�ŏ㏑��
             myArray2(m, 1) = myArray2(m - 1, 1)
            counterTest = counterTest + 1
            'If���̌��ʂ�TRUE�ł������ꍇ�A�J�E���^�[�ɂP�������AtestArray�ɃC���N�������^���ɒl���i�[���Ă���
            t = t + 1
        End If
    Next o
Next m
Debug.Print UBound(myArray2)

'�d���폜��M�\����O�\�f�[�^�m�F�̂��߁A�V�[�g��V�K�쐬�h
Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "M�\����O�\"
'�z��̏d�����폜���邽�߂Ɉ�x�V�[�g�ɏo��
Worksheets("M�\����O�\").range("B1:B" & UBound(myArray2, 1)) = myArray2
'�V�[�g�ɏo�͂��ꂽ�l�̏d�����폜
Worksheets("M�\����O�\").range("B1:B" & UBound(myArray2, 1)).RemoveDuplicates Columns:=1, Header:=xlNo
'�d�����폜�����f�[�^��z��ɍĊi�[
Erase myArray2
Worksheets("M�\����O�\").Select
lastRow2 = Worksheets("M�\����O�\").Cells(1, 2).End(xlDown).Row '�ŏI�s�ԍ��̎擾
Set range2 = Worksheets("M�\����O�\").range(Cells(1, 2), Cells(lastRow2, 2)) '�z��Ɋi�[����l�̃Z���͈͂�I��
myArray2 = (range2) '�z��Ɋi�[
'M�\(���)����O�\(���)�̑Ώێ҂��o��
Worksheets("M�\����O�\").range("A1:A" & UBound(testArray)) = testArray
Debug.Print UBound(myArray2)
Debug.Print counterTest

'M�\(���)�����p�ɕϐ���������
t = 1 '��ԍ��̏�����
counterTest = 0 '�d���J�E���^�[�̏�����
Erase testArray '�z��̏�����

ReDim testArray(1 To 10000, 1 To 1)

'M�\�i���j����O�\(���)�������z��(M-O)���쐬�B�i�v�f���͍폜�ł��Ȃ����߁A����ς݂̑O�̗v�f�ŏ㏑���j

For m = LBound(myArray5, 1) To UBound(myArray5, 1)
    For o = LBound(myArray3, 1) To UBound(myArray3, 1)
        If myArray5(m, 1) = myArray3(o, 1) And myArray4(o, 1) = "��앪" Then
            'M�\�i��Áj����O�\�̃f�[�^��z��Ɋi�[
            testArray(t, 1) = myArray5(m, 1)
            '����ς̗v�f�ŏ㏑��
             myArray5(m, 1) = myArray5(m - 1, 1)
            counterTest = counterTest + 1
            'If���̌��ʂ�TRUE�ł������ꍇ�A�J�E���^�[�ɂP�������AtestArray�ɃC���N�������^���ɒl���i�[���Ă���
            t = t + 1
        End If
    Next o
Next m
Debug.Print UBound(myArray5)

'�z��̏d�����폜���邽�߂Ɉ�x�V�[�g�ɏo��
Worksheets("M�\����O�\").range("D1:D" & UBound(myArray5, 1)) = myArray5
'�V�[�g�ɏo�͂��ꂽ�l�̏d�����폜
Worksheets("M�\����O�\").range("D1:D" & UBound(myArray5, 1)).RemoveDuplicates Columns:=1, Header:=xlNo
'�d�����폜�����f�[�^��z��ɍĊi�[
Erase myArray5 '�z��̏�����
Worksheets("M�\����O�\").Select
lastRow5 = Worksheets("M�\����O�\").Cells(1, 4).End(xlDown).Row '�ŏI�s�ԍ��̎擾
Set range5 = Worksheets("M�\����O�\").range(Cells(1, 4), Cells(lastRow5, 4)) '�z��Ɋi�[����l�̃Z���͈͂�I��
myArray5 = (range5) '�z��Ɋi�[
'M�\(���)����O�\(���)�̑Ώێ҂��o��
Worksheets("M�\����O�\").range("C1:C" & UBound(testArray)) = testArray
Debug.Print UBound(myArray5)
Debug.Print counterTest
'testArray�̏�����
Erase testArray
'O�\�֘A�̔z����폜
Erase myArray3
Erase myArray4

'M�\����PQ�\�W�v�����͂�������----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim myArray6 As Variant 'P�\�f�[�^�i�[�p�z��i������j
Dim myArray7 As Variant 'P�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj
Dim myArray8 As Variant 'Q�\�f�[�^�i�[�p�z��i������j
Dim myArray9 As Variant 'Q�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj

Dim lastRow6 As Long 'P�\�f�[�^�i�[�p�z��i������j�ŏI�s�ԍ�
Dim lastRow7 As Long 'P�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj�ŏI�s�ԍ�
Dim lastRow8 As Long 'Q�\�f�[�^�i�[�p�z��i������j�ŏI�s�ԍ�
Dim lastRow9 As Long 'Q�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj�ŏI�s�ԍ�

Dim range6 As range 'P�\�f�[�^�z��i�[�͈́i������j
Dim range7 As range 'P�\�f�[�^�z��i�[�͈́i�ی��Łm���n��ʁj
Dim range8 As range 'Q�\�f�[�^�z��i�[�͈́i������j
Dim range9 As range 'Q�\�f�[�^�z��i�[�͈́i�ی��Łm���n��ʁj

'�͈͍ŏI�s�ԍ��̎擾/�z��i�[�f�[�^�͈͂̎擾

'P�\�i�����j
ThisWorkbook.Worksheets(WorksheetIndexP).Select
lastRow6 = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(1, 1).End(xlDown).Row
Set range6 = ThisWorkbook.Worksheets(WorksheetIndexP).range(Cells(2, standardAtenaP), Cells(lastRow6, standardAtenaP))

'P�\�i�ی��Łm���n��ʁj
lastRow7 = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(1, 1).End(xlDown).Row
Set range7 = ThisWorkbook.Worksheets(WorksheetIndexP).range(Cells(2, standardSyubetuP), Cells(lastRow7, standardSyubetuP))

'O�\�i�����j
ThisWorkbook.Worksheets(WorksheetIndexQ).Select
lastRow8 = ThisWorkbook.Worksheets(WorksheetIndexQ).Cells(1, 1).End(xlDown).Row
Set range8 = ThisWorkbook.Worksheets(WorksheetIndexQ).range(Cells(2, standardAtenaQ), Cells(lastRow8, standardAtenaQ))

'O�\�i�ی��Łm���n��ʁj
lastRow9 = ThisWorkbook.Worksheets(WorksheetIndexQ).Cells(1, 1).End(xlDown).Row
Set range9 = ThisWorkbook.Worksheets(WorksheetIndexQ).range(Cells(2, standardSyubetuQ), Cells(lastRow9, standardSyubetuQ))


'�e�z��(�Q�����z��)�Ƀf�[�^���i�[
myArray6 = (range6) 'P�\(����)�f�[�^�z��
myArray7 = (range7) 'P�\(���)�f�[�^�z��
myArray8 = (range8) 'Q�\(����)�f�[�^�z��
myArray9 = (range9) 'Q�\(���)�f�[�^�z��

'�J�ԏ�������p
Dim pq As Long 'PQ�\�p

'(M-O)�\(���)����P�\(���)�W�v�p�J�E���^�[
Dim counterMPI As Long
counterMPI = 0

'(M-O)�\(���)����P�\(���)�W�v�p�J�E���^�[
Dim counterMPK As Long
counterMPK = 0

'(M-O)�\(���)����Q�\(���)�W�v�p�J�E���^�[
Dim counterMQI As Long
counterMQI = 0

'(M-O)�\(���)����Q�\(���)�W�v�p�J�E���^�[
Dim counterMQK As Long
counterMQK = 0


'(M-O)�\(���)����P�\(���)�W�v
For m = LBound(myArray2, 1) To UBound(myArray2, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray2(m, 1) = myArray6(pq, 1) And myArray7(pq, 1) = "��Õ�" Then
            counterMPI = counterMPI + 1
        End If
    Next pq
Next m
Debug.Print counterMPI


'(M-O)�\(���)����P�\(���)�W�v
For m = LBound(myArray5, 1) To UBound(myArray5, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray5(m, 1) = myArray6(pq, 1) And myArray7(pq, 1) = "��앪" Then
            counterMPK = counterMPK + 1
        End If
    Next pq
Next m
Debug.Print counterMPK


'(M-O)�\(���)����Q�\(���)�W�v
For m = LBound(myArray2, 1) To UBound(myArray2, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray2(m, 1) = myArray8(pq, 1) And myArray9(pq, 1) = "��Õ�" Then
            counterMQI = counterMQI + 1
        End If
    Next pq
Next m
Debug.Print counterMQI


'(M-O)�\(���)����Q�\(���)�W�v
For m = LBound(myArray5, 1) To UBound(myArray5, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray5(m, 1) = myArray8(pq, 1) And myArray9(pq, 1) = "��앪" Then
            counterMQK = counterMQK + 1
        End If
    Next pq
Next m
Debug.Print counterMQK


'M�\�֘A�̔z���������
Erase myArray2 'PQ�\�̏W�v���I���܂Ŏc��
Erase myArray5 'PQ�\�̏W�v���I���܂Ŏc��

'N�\����PQ�\�W�v�����͂�������----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim myArray10 As Variant 'N�\�f�[�^�i�[�p�z��i������j
Dim myArray11 As Variant 'N�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj

Dim lastRow10 As Long 'N�\�f�[�^�i�[�p�z��i������j�ŏI�s�ԍ�
Dim lastRow11 As Long 'N�\�f�[�^�i�[�p�z��i�ی��Łm���n��ʁj�ŏI�s�ԍ�

Dim range10 As range 'N�\�f�[�^�z��i�[�͈́i������j
Dim range11 As range 'N�\�f�[�^�z��i�[�͈́i�ی��Łm���n��ʁj

'�J�ԏ�������p
Dim n As Long 'N�\�p

'N�\(���)����P�\(���)�W�v�p�J�E���^�[
Dim counterNPI As Long
counterNPI = 0

'N�\(���)����P�\(���)�W�v�p�J�E���^�[
Dim counterNPK As Long
counterNPK = 0

'N�\(���)����Q�\(���)�W�v�p�J�E���^�[
Dim counterNQI As Long
counterNQI = 0

'N�\(���)����Q�\(���)�W�v�p�J�E���^�[
Dim counterNQK As Long
counterNQK = 0


'�͈͍ŏI�s�ԍ��̎擾/�z��i�[�f�[�^�͈͂̎擾

'N�\�i�����j
ThisWorkbook.Worksheets(WorksheetIndexN).Select
lastRow10 = ThisWorkbook.Worksheets(WorksheetIndexN).Cells(1, 1).End(xlDown).Row
Set range10 = ThisWorkbook.Worksheets(WorksheetIndexN).range(Cells(2, standardAtenaN), Cells(lastRow10, standardAtenaN))

'N�\�i�ی��Łm���n��ʁj
lastRow11 = ThisWorkbook.Worksheets(WorksheetIndexN).Cells(1, 1).End(xlDown).Row
Set range11 = ThisWorkbook.Worksheets(WorksheetIndexN).range(Cells(2, standardSyubetuN), Cells(lastRow11, standardSyubetuN))

'�e�z��(�Q�����z��)�Ƀf�[�^���i�[
myArray10 = (range10) 'N�\(����)�f�[�^�z��
myArray11 = (range11) 'N�\(���)�f�[�^�z��

'N�\(���)����P�\(���)�W�v
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray10(n, 1) = myArray6(pq, 1) And myArray11(n, 1) = "��Õ�" And myArray7(pq, 1) = "��Õ�" Then
            counterNPI = counterNPI + 1
        End If
    Next pq
Next n
Debug.Print counterNPI

'N�\(���)����P�\(���)�W�v
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray6, 1) To UBound(myArray6, 1)
        If myArray10(n, 1) = myArray6(pq, 1) And myArray11(n, 1) = "��앪" And myArray7(pq, 1) = "��앪" Then
            counterNPK = counterNPK + 1
        End If
    Next pq
Next n
Debug.Print counterNPK

'N�\(���)����Q�\(���)�W�v
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray10(n, 1) = myArray8(pq, 1) And myArray11(n, 1) = "��Õ�" And myArray9(pq, 1) = "��Õ�" Then
            counterNQI = counterNQI + 1
        End If
    Next pq
Next n
Debug.Print counterNQI

'N�\(���)����Q�\(���)�W�v
For n = LBound(myArray10, 1) To UBound(myArray10, 1)
    For pq = LBound(myArray8, 1) To UBound(myArray8, 1)
        If myArray10(n, 1) = myArray8(pq, 1) And myArray11(n, 1) = "��앪" And myArray9(pq, 1) = "��앪" Then
            counterNQK = counterNQK + 1
        End If
    Next pq
Next n
Debug.Print counterNQK

'PQN�\�̔z���������
Erase myArray6
Erase myArray7
Erase myArray8
Erase myArray9
Erase myArray10
Erase myArray11

'���ʂ̏o�͂͂�������----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

ThisWorkbook.Worksheets(1).Select
ThisWorkbook.Worksheets(1).range("D20").Value = counterMPI '(M-O)�\(���)����P�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("G20").Value = counterMPK '(M-O)�\(���)����P�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("D23").Value = counterMQI '(M-O)�\(���)����Q�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("G23").Value = counterMQK '(M-O)�\(���)����Q�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("D21").Value = counterNPI 'N�\(���)����P�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("G21").Value = counterNPK 'N�\(���)����P�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("D24").Value = counterNQI 'N�\(���)����P�\(���)�W�v���ʂ��o��
ThisWorkbook.Worksheets(1).range("G24").Value = counterNQK 'N�\(���)����P�\(���)�W�v���ʂ��o��
End Sub
