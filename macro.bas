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
        '振込表のB列に会社が存在する場合
        If Not Cells(j, "B").Value = "" Then
            '会社名出力
            Debug.Print Cells(j, "B").Value
            Get_First_Month (Cells(j, "B").Value)
        '振込表のB列に会社が存在しない場合
        Else
            Debug.Print "振込表の会社名が空白です。"
        End If
    Next
    '処理終了コメント
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
    'シートが存在する場合
    If Flag = 1 Then
    
        Debug.Print "シートは存在します。"
        
        'A列の2行目が空白セルでない場合
        If Not Worksheets(name).Cells(2, "A") = "" Then
            mon = Month(Worksheets(name).Cells(2, "A"))
            loading (2), (mon), (name)
        'A列の3行目が空白セルの場合
        ElseIf Worksheets(name).Cells(3, "A") = "" Then
            MsgBox name & "の日付が入力されていません。"
        'A列の3行目が空白セルでない場合
        Else
            mon = Month(Worksheets(name).Cells(3, "A"))
            loading (3), (mon), (name)
        End If
    'シートが存在しない場合
    Else
        MsgBox name & "のシートが存在しません。"
        Debug.Print "シートは存在しません。"
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
        '月が変わらない場合
        If mon = temp Then
            num = num + 1
            mon = temp
      
        Else
            Debug.Print mon & "月"
            Debug.Print "個数 : " & num
            Debug.Print "i : " & i
            num = 0
            
            money = Get_Money(name, i - 1)
            
            Debug.Print "money : " & money
            
            column_num = Search_Row(name)
            
            monID = Get_Month_Cell(mon)
            
            If Not monID = "" Then
                Worksheets("振込表").Cells(column_num, monID) = money
            End If
            
            mon = Month(Worksheets(name).Cells(i, "A"))

            loading (i), (mon), (name)
        End If
    Next
    
End Sub
'振込表の月のセルを取得
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
        
        Worksheets("振込表").Cells(column_num, "Z") = money

    End If
    If Day(Worksheets(name).Cells(i, "A")) = 31 Then

        money = Get_Money(name, i)
        column_num = Search_Row(name)
        
        money = money - Worksheets("振込表").Cells(column_num, "Z")
        Worksheets("振込表").Cells(column_num, "AA") = money
        
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

    For A = 2 To Worksheets("振込表").Cells(Rows.Count, 1).End(xlUp).Row + 2
            
        '振込表で対象の会社が記載されている行を取得
        If Worksheets("振込表").Cells(A, "B") = name Then
            Search_Row = A
        End If
    Next
End Function
