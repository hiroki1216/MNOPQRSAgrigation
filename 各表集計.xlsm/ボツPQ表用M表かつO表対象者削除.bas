Attribute VB_Name = "�{�cPQ�\�pM�\����O�\�Ώێҍ폜"
Sub PQ�\�pM�\����O�\�Ώێҍ폜()

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

'���[�N�V�[�g���ċA�I�ɏ���
For Each Worksheet In Worksheets
    '�Ώۃ��[�N�V�[�g�̈����ԍ��J�����̗�ԍ��ƃ��[�N�V�[�g�̃C���f�b�N�X�ԍ��̎擾
    If Worksheet.Name Like "*M�\*" & "*���*" Then
        WorksheetIndexMI = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�����ԍ�" Then
                standardAtenaMI = i
                Worksheet.Cells(1, standardAtenaMI).Interior.ColorIndex = 3
            End If
        Next i
    ElseIf Worksheet.Name Like "*M�\*" & "*���*" Then
        WorksheetIndexMK = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�����ԍ�" Then
                standardAtenaMK = i
                Worksheet.Cells(1, standardAtenaMK).Interior.ColorIndex = 3
            End If
        Next i
        
    ElseIf Worksheet.Name Like "*N�\*" Then
        WorksheetIndexN = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�����ԍ�" Then
                standardAtenaN = i
                Worksheet.Cells(1, standardAtenaN).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�ی��Łm���n���" Then
                standardSyubetuN = i
                Worksheet.Cells(1, standardSyubetuN).Interior.ColorIndex = 3
            End If
        Next i
        
     ElseIf Worksheet.Name Like "*O�\*" Then
        WorksheetIndexO = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�����ԍ�" Then
                standardAtenaO = i
                Worksheet.Cells(1, standardAtenaO).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�ی��Łm���n���" Then
                standardSyubetuO = i
                Worksheet.Cells(1, standardSyubetuO).Interior.ColorIndex = 3
            End If
        Next i
    
    ElseIf Worksheet.Name Like "*P�\*" Then
        WorksheetIndexP = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�����ԍ�" Then
                standardAtenaP = i
                Worksheet.Cells(1, standardAtenaP).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�ی��Łm���n���" Then
                standardSyubetuP = i
                Worksheet.Cells(1, standardSyubetuP).Interior.ColorIndex = 3
            End If
        Next i
        
    ElseIf Worksheet.Name Like "*Q�\*" Then
        WorksheetIndexQ = Worksheet.Index
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�����ԍ�" Then
                standardAtenaQ = i
                Worksheet.Cells(1, standardAtenaQ).Interior.ColorIndex = 3
            End If
        Next i
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "�ی��Łm���n���" Then
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


'M�\����O�\�Ώێ҂��폜�͂�������
Dim m As Long 'M�\�J�ԏ�������p
Dim o As Long 'O�\�J�ԏ�������p
Dim lastRowMI As Long 'M�\(���)�J�ԏ�������p
Dim lastRowMK As Long 'M�\(���)�J�ԏ�������p
Dim lastRowO As Long 'O�\�J�ԏ�������p


lastRowMI = ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(1, 1).End(xlDown).Row

lastRowMK = ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(1, 1).End(xlDown).Row

lastRowO = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(1, 1).End(xlDown).Row



'M�\(��Õ�)����O�\����
For o = lastRowO To 2 Step -1
    For m = lastRowMI To 2 Step -1
        If ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).Value = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardAtenaO).Value And ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardSyubetuO).Value = "��Õ�" Then
            'ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).EntireRow.Delete
            ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).Interior.ColorIndex = 3
        End If
    Next m
Next o
'M�\(��앪)����O�\����
For o = lastRowO To 2 Step -1
    For m = lastRowMK To 2 Step -1
        If ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(m, standardAtenaMK).Value = ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardAtenaO).Value And ThisWorkbook.Worksheets(WorksheetIndexO).Cells(o, standardSyubetuO).Value = "��앪" Then
            'ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(m, standardAtenaMK).EntireRow.Delete
            ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(m, standardAtenaMK).Interior.ColorIndex = 3
        End If
    Next m
Next o

'PQ�\�W�v�͂�������
'Dim counterP As Long 'P�\��ی��Ґ��J�E���^�[
'Dim counterQ As Long 'Q�\��ی��Ґ��J�E���^�[
'Dim pq As Long 'PQ�\�J�ԏ�������p
'Dim n As Long 'N�J�ԏ�������p
'Dim lastRowP As Long 'P�\�J�ԏ�������p
'Dim lastRowQ As Long 'Q�\�J�ԏ�������p
'Dim lastRowN As Long 'N�\�J�ԏ�������p
'
'lastRowP = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(1, 1).End(xlDown).Row
'lastRowQ = ThisWorkbook.Worksheets(WorksheetIndexQ).Cells(1, 1).End(xlDown).Row
'lastRowN = ThisWorkbook.Worksheets(WorksheetIndexN).Cells(1, 1).End(xlDown).Row
'lastRowMI = ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(1, 1).End(xlDown).Row
'lastRowMK = ThisWorkbook.Worksheets(WorksheetIndexMK).Cells(1, 1).End(xlDown).Row

'P�\�W�v(��Õ�)
'counterP = 0
'For m = 2 To lastRowMI
'    For pq = 2 To lastRowP
'        If ThisWorkbook.Worksheets(WorksheetIndexMI).Cells(m, standardAtenaMI).Value = ThisWorkbook.Worksheets(WorksheetIndexP).Cells(pq, standardAtenaP).Value And ThisWorkbook.Worksheets(WorksheetIndexP).Cells(pq, standardSyubetuP).Value = "��Õ�" And ThisWorkbook.Worksheets(WorksheetIndexP).Cells(pq, standardSyubetuP).Value > 0 Then
'            counterP = counterP + 1
'        End If
'    Next pq
'Next m
'Debug.Print counterP
End Sub
