Attribute VB_Name = "PQ�\�p�s�v�f�[�^�̍폜"
Sub PQ�\�p�s�v�f�[�^�̍폜()
    Dim Worksheet As Worksheet '���[�N�V�[�g�擾�p
    Dim i As Long '�J�Ԃ������p
    Dim lastcolumn As Long '�J�ԏ����I�_�p
    Dim standardRecord As Long '���R�[�h�敪�̗�ԍ��p
    Dim standardGinushhi As Long '�[��敪�̗�ԍ��p
    Dim standardTokuteisya As Long '���蓯�ꐢ�я����ҋ敪�̗�ԍ��p
    Dim standardKijyunnsousyotoku As Long '��������i��~�����؎́j�̗�ԍ��p

    
    '�e�V�[�g�̃J���������������鏈��
    For Each Worksheet In Worksheets
        lastcolumn = Worksheet.Cells(1, 1).End(xlToRight).Column
        '�J�����̌���
        For i = 1 To lastcolumn
            If Worksheet.Cells(1, i).Value = "���R�[�h�敪" Then
                standardRecord = i
                Worksheet.Cells(1, standardRecord).Interior.ColorIndex = 3
                'Debug.Print (standardRecord)
            ElseIf Worksheet.Cells(1, i).Value = "�[��敪" Then
                standardGinushhi = i
                Worksheet.Cells(1, standardGinushhi).Interior.ColorIndex = 3
                'Debug.Print (standardGinushhi)
            ElseIf Worksheet.Cells(1, i).Value = "���蓯�ꐢ�я����ҋ敪" Then
                standardTokuteisya = i
                Worksheet.Cells(1, standardTokuteisya).Interior.ColorIndex = 3
                'Debug.Print (standardTokuteisya)
            ElseIf Worksheet.Cells(1, i).Value = "��������i��~�����؎́j" Then
            standardKijyunnsousyotoku = i
            Worksheet.Cells(1, standardKijyunnsousyotoku).Interior.ColorIndex = 3
            'Debug.Print (standardTokuteisya)
            End If
        Next i
        
        '�s�v�s�̍폜�����͂�������
        
        Dim s As Long '�J�ԏ����p
        Dim lastRow As Long '�����V�[�g�̍ŏI�s
        'A��̍ŏI�s���擾
        lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        'standardRecord=0�̏ꍇ�́A�������X�L�b�v����
        If standardRecord = 0 Then
            GoTo Continue
        End If
        '�s�v�s�̍폜����
        '���R�[�h�敪
        For s = lastRow To 2 Step -1
            If Worksheet.Cells(s, standardRecord).Value = "����" Then
                Worksheet.Cells(s, standardRecord).EntireRow.Delete
            End If
        Next s
        
        '�[�����ю�敪
        lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
        For s = lastRow To 2 Step -1
            If Worksheet.Cells(s, standardGinushhi).Value = "�[�����ю�" Then
                Worksheet.Cells(s, standardGinushhi).EntireRow.Delete
            End If
        Next s
        
        '���蓯�ꐢ�я����ҋ敪
        lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
        For s = lastRow To 2 Step -1
            If Worksheet.Cells(s, standardTokuteisya).Value = "���蓯�ꐢ�я�����" Then
                Worksheet.Cells(s, standardTokuteisya).EntireRow.Delete
            End If
        Next s
        '��������i��~�����؎́jPQ�\�̂ݏ�������
        If Worksheet.Name Like "*P�\*" Then
            lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
            For s = lastRow To 2 Step -1
                If Worksheet.Cells(s, standardKijyunnsousyotoku).Value = 0 Then
                    Worksheet.Cells(s, standardKijyunnsousyotoku).EntireRow.Delete
                End If
            Next s
        ElseIf Worksheet.Name Like "*Q�\*" Then
            lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
            For s = lastRow To 2 Step -1
                If Worksheet.Cells(s, standardKijyunnsousyotoku).Value = 0 Then
                    Worksheet.Cells(s, standardKijyunnsousyotoku).EntireRow.Delete
                End If
            Next s
        Else
            GoTo Continue
        End If
        
    
Continue:
    Next
End Sub

