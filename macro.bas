Attribute VB_Name = "macro"
'会社を探す関数
Sub Search_Company()
    
    Dim lastrow
    
    For i = 2 To Worksheets("振込表").Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(i, "B").Value = "小村分店振込" Then
            lastrow = i
        End If
    Next
    
    For j = 2 To lastrow - 1
        If Not Cells(j, "B").Value = "" Then
            '会社名出力
            Debug.Print Cells(j, "B").Value
            Get_First_Month (Cells(j, "B").Value)
        Else
          Debug.Print "振込表の会社名が空白です。"
        End If
    Next
    MsgBox "処理完了"
End Sub
'各会社のシートで初めの月を取得する関数
Sub Get_First_Month(name As String)

    Dim mon
    Dim Flag
    Flag = 0 '初期値
    
    'Sheetの存在をチェック
    For i = 1 To Sheets.Count
        If Sheets(i).name = name Then
            Flag = 1
        End If
    Next
    If Flag = 1 Then
        Debug.Print "シートは存在します。"
        
        If Not Worksheets(name).Cells(2, "A") = "" Then
            mon = Month(Worksheets(name).Cells(2, "A"))
            syori (2), (mon), (name)
        ElseIf Worksheets(name).Cells(3, "A") = "" Then
            MsgBox name & "の日付が入力されていません。"
        Else
            mon = Month(Worksheets(name).Cells(3, "A"))
            syori (3), (mon), (name)
        
        End If
    Else
        MsgBox name & "のシートが存在しません。"
        Debug.Print "シートは存在しません。"
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
            Debug.Print mon & "月"
            Debug.Print "個数 : " & num
            Debug.Print "i : " & i
            num = 0
            
            money = Worksheets(name).Cells(i - 1, "I")
            If money = 0 Then
                money = ""
            End If
            
            Debug.Print "money : " & money
            
            For a = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 2
                        
                If Worksheets("振込表").Cells(a, "B") = name Then
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
            Worksheets("振込表").Cells(column_num, monID) = money
            
            mon = Month(Worksheets(name).Cells(i, "A"))
            syori (i), (mon), (name)
        End If
    Next
    
End Sub



