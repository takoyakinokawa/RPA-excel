Attribute VB_Name = "macro"
'��Ђ�T���֐�
Sub Search_Company()
    Dim lastrow
    
    For i = 2 To Worksheets("�U���\").Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(i, "B").Value = "�������X�U��" Then
            lastrow = i
        End If
    Next
    
    For j = 2 To lastrow - 1
        '�U���\��B��ɉ�Ђ����݂���ꍇ
        If Not Cells(j, "B").Value = "" Then
            '��Ж��o��
            Debug.Print Cells(j, "B").Value
            Get_First_Month (Cells(j, "B").Value)
        '�U���\��B��ɉ�Ђ����݂��Ȃ��ꍇ
        Else
            Debug.Print "�U���\�̉�Ж����󔒂ł��B"
        End If
    Next
    '�����I���R�����g
    MsgBox "��������"
    
End Sub
'�e��Ђ̃V�[�g�ŏ��߂̌����擾����֐�
Sub Get_First_Month(name As String)

    Dim mon
    Dim Flag
    Flag = 0 '�����l
    
    'Sheet�̑��݂��`�F�b�N
    For i = 1 To Sheets.Count
        If Sheets(i).name = name Then
            Flag = 1
        End If
    Next
    '�V�[�g�����݂���ꍇ
    If Flag = 1 Then
    
        Debug.Print "�V�[�g�͑��݂��܂��B"
        
        'A���2�s�ڂ��󔒃Z���łȂ��ꍇ
        If Not Worksheets(name).Cells(2, "A") = "" Then
            mon = Month(Worksheets(name).Cells(2, "A"))
            loading (2), (mon), (name)
        'A���3�s�ڂ��󔒃Z���̏ꍇ
        ElseIf Worksheets(name).Cells(3, "A") = "" Then
            MsgBox name & "�̓��t�����͂���Ă��܂���B"
        'A���3�s�ڂ��󔒃Z���łȂ��ꍇ
        Else
            mon = Month(Worksheets(name).Cells(3, "A"))
            loading (3), (mon), (name)
        End If
    '�V�[�g�����݂��Ȃ��ꍇ
    Else
        MsgBox name & "�̃V�[�g�����݂��܂���B"
        Debug.Print "�V�[�g�͑��݂��܂���B"
    End If
    
End Sub
Sub loading(first_column As Integer, mon As Integer, name As String)
    
    Dim temp
    
    temp = Worksheets(name).Cells(Rows.Count, 1).End(xlUp).Row
    For i = first_column To Worksheets(name).Cells(temp, 1).End(xlUp).Row + 1
        temp = Month(Worksheets(name).Cells(i, "A"))
        
        If mon = 3 Then
            March (mon), (name), (i)
        End If
        '�����ς��Ȃ��ꍇ
        If mon = temp Then
            num = num + 1
            mon = temp
      
        Else
            Debug.Print mon & "��"
            Debug.Print "�� : " & num
            Debug.Print "i : " & i
            num = 0
            
            money = Get_Money(name, i - 1)
            
            Debug.Print "money : " & money
            
            column_num = Search_Row(name)
            
            monID = Get_Month_Cell(mon)
            
            If Not monID = "" Then
                Worksheets("�U���\").Cells(column_num, monID) = money
            End If
            
            mon = Month(Worksheets(name).Cells(i, "A"))

            loading (i), (mon), (name)
        End If
    Next
    
End Sub
'�U���\�̌��̃Z�����擾
Function Get_Month_Cell(mon) As String

    If mon = 4 Then
        monID = "D"
    End If
    If mon = 5 Then
        monID = "F"
    End If
    If mon = 6 Then
        monID = "H"
    End If
    If mon = 7 Then
        monID = "J"
    End If
    If mon = 8 Then
        monID = "L"
    End If
    If mon = 9 Then
        monID = "N"
    End If
    If mon = 10 Then
        monID = "P"
    End If
    If mon = 11 Then
        monID = "R"
    End If
    If mon = 12 Then
        monID = "T"
    End If
    If mon = 1 Then
        monID = "V"
    End If
    If mon = 2 Then
        monID = "X"
    End If
    
    Get_Month_Cell = monID
    
End Function
Sub March(mon, name, i)
    
    init = 20
    
    If Day(Worksheets(name).Cells(i, "A")) = init Then

        money = Get_Money(name, i)
        column_num = Search_Row(name)
        
        Worksheets("�U���\").Cells(column_num, "Z") = money

    End If
    If Day(Worksheets(name).Cells(i, "A")) = 31 Then

        money = Get_Money(name, i)
        column_num = Search_Row(name)
        
        money = money - Worksheets("�U���\").Cells(column_num, "Z")
        Worksheets("�U���\").Cells(column_num, "AA") = money
        
    End If
    
End Sub
Function Get_Money(name, i)
    money = Worksheets(name).Cells(i, "I")
    If money = 0 Then
        money = ""
    End If
    Debug.Print "money : " & money
    Get_Money = money
End Function
Function Search_Row(name)

    For A = 2 To Worksheets("�U���\").Cells(Rows.Count, 1).End(xlUp).Row + 2
            
        '�U���\�őΏۂ̉�Ђ��L�ڂ���Ă���s���擾
        If Worksheets("�U���\").Cells(A, "B") = name Then
            Search_Row = A
        End If
    Next
End Function
