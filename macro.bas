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
        If Not Cells(j, "B").Value = "" Then
            '��Ж��o��
            Debug.Print Cells(j, "B").Value
            Get_First_Month (Cells(j, "B").Value)
        Else
          Debug.Print "�U���\�̉�Ж����󔒂ł��B"
        End If
    Next
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
    If Flag = 1 Then
        Debug.Print "�V�[�g�͑��݂��܂��B"
        
        If Not Worksheets(name).Cells(2, "A") = "" Then
            mon = Month(Worksheets(name).Cells(2, "A"))
            syori (2), (mon), (name)
        ElseIf Worksheets(name).Cells(3, "A") = "" Then
            MsgBox name & "�̓��t�����͂���Ă��܂���B"
        Else
            mon = Month(Worksheets(name).Cells(3, "A"))
            syori (3), (mon), (name)
        
        End If
    Else
        MsgBox name & "�̃V�[�g�����݂��܂���B"
        Debug.Print "�V�[�g�͑��݂��܂���B"
    End If
    
End Sub
Sub syori(first_column As Integer, mon As Integer, name As String)
    
    Dim temp
    
    temp = Worksheets(name).Cells(Rows.Count, 1).End(xlUp).Row
    For i = first_column To Worksheets(name).Cells(temp, 1).End(xlUp).Row + 1
        temp = Month(Worksheets(name).Cells(i, "A"))
        
        If mon = temp Then
            num = num + 1
            mon = temp
      
        Else
            Debug.Print mon & "��"
            Debug.Print "�� : " & num
            Debug.Print "i : " & i
            num = 0
            
            money = Worksheets(name).Cells(i - 1, "I")
            If money = 0 Then
                money = ""
            End If
            
            Debug.Print "money : " & money
            
            For a = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 2
                        
                If Worksheets("�U���\").Cells(a, "B") = name Then
                    column_num = a
                End If
            Next
            
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
            Worksheets("�U���\").Cells(column_num, monID) = money
            
            mon = Month(Worksheets(name).Cells(i, "A"))
            syori (i), (mon), (name)
        End If
    Next
    
End Sub


