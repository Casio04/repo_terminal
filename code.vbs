Sub excel_challenge()
    
    Application.ScreenUpdating = False
    
    Dim hw As String
    Dim lastrow, last_unique, stock_count As Integer
    Dim CloseValue, OpenValue As Double
    hw = ActiveWorkbook.Name
    
    
    For Sheet = 1 To Sheets.Count
        lastrow = Workbooks(hw).Sheets(Sheet).Cells(Rows.Count, "A").End(xlUp).Row
    
        'Naming headers
        Workbooks(hw).Sheets(Sheet).Range("I1").Value = "Ticker"
        Workbooks(hw).Sheets(Sheet).Range("J1").Value = "Yearly Change"
        Workbooks(hw).Sheets(Sheet).Range("K1").Value = "Percent Change"
        Workbooks(hw).Sheets(Sheet).Range("L1").Value = "Total Stock Volume@"
        
        'Removing all duplicates from column A and inserting unique values into I
        Workbooks(hw).Sheets(Sheet).Range("A2:A" & lastrow).Copy Workbooks(hw).Sheets(Sheet).Range("I2")
        Workbooks(hw).Sheets(Sheet).Range("I1:I" & lastrow).RemoveDuplicates Columns:=1, Header:=xlYes
        
        last_unique = Workbooks(hw).Sheets(Sheet).Cells(Rows.Count, "I").End(xlUp).Row
        
        'Get the first and last date by stock and fill Yearly Change column
        OpenValue = Workbooks(hw).Sheets(Sheet).Range("C2").Value
        stock_count = 2
        For i = 2 To lastrow
                If Workbooks(hw).Sheets(Sheet).Range("A" & i).Value = Workbooks(hw).Sheets(Sheet).Range("A" & i + 1).Value Then
                Else
                CloseValue = Workbooks(hw).Sheets(Sheet).Range("F" & i).Value
                Workbooks(hw).Sheets(Sheet).Range("J" & stock_count).Value = CloseValue - OpenValue
                'Format Yearly Change
                    If Workbooks(hw).Sheets(Sheet).Range("J" & stock_count).Value > 0 Then
                        Workbooks(hw).Sheets(Sheet).Range("J" & stock_count).Interior.ColorIndex = 4
                    Else
                        Workbooks(hw).Sheets(Sheet).Range("J" & stock_count).Interior.ColorIndex = 3
                    End If
                'Fill Percent change considering possibility of a 0
                    If OpenValue = 0 Then
                        Workbooks(hw).Sheets(Sheet).Range("K" & stock_count).Value = 0
                    Else
                        Workbooks(hw).Sheets(Sheet).Range("K" & stock_count).Value = CloseValue / OpenValue
                    End If
                Workbooks(hw).Sheets(Sheet).Range("L" & stock_count).Formula = "=SUMIFS(G2:G" & lastrow & ",A2:A" & lastrow & ",I" & stock_count & ")"
                Workbooks(hw).Sheets(Sheet).Range("L" & stock_count).Formula = Workbooks(hw).Sheets(Sheet).Range("L" & stock_count).Value
                stock_count = stock_count + 1
                OpenValue = Workbooks(hw).Sheets(Sheet).Range("C" & i + 1).Value
                End If
        Next i
        
        'Challenge
        Workbooks(hw).Sheets(Sheet).Range("O2") = "Greatest % Increase"
        Workbooks(hw).Sheets(Sheet).Range("O3") = "Greatest % Decrease"
        Workbooks(hw).Sheets(Sheet).Range("O4") = "Greatest Total Volume"
        Workbooks(hw).Sheets(Sheet).Range("P1") = "Ticker"
        Workbooks(hw).Sheets(Sheet).Range("Q1") = "Value"
        Workbooks(hw).Sheets(Sheet).Range("Q2").Formula = "=MAX(K2:K" & last_unique & ")"
        Workbooks(hw).Sheets(Sheet).Range("P2").Formula = "=INDEX(I2:I" & last_unique & ",MATCH(Q2,K2:K" & last_unique & ",0),0)"
        Workbooks(hw).Sheets(Sheet).Range("Q3").Formula = "=MIN(K2:K" & last_unique & ")"
        Workbooks(hw).Sheets(Sheet).Range("P3").Formula = "=INDEX(I2:I" & last_unique & ",MATCH(Q3,K2:K" & last_unique & ",0),0)"
        Workbooks(hw).Sheets(Sheet).Range("Q4").Formula = "=MAX(L2:L" & last_unique & ")"
        Workbooks(hw).Sheets(Sheet).Range("P4").Formula = "=INDEX(I2:I" & last_unique & ",MATCH(Q4,L2:L" & last_unique & ",0),0)"
        
        'Format cells
        Workbooks(hw).Sheets(Sheet).Columns.AutoFit
        Workbooks(hw).Sheets(Sheet).Columns.HorizontalAlignment = xlCenter
        Workbooks(hw).Sheets(Sheet).Range("K2:K" & lastrow).NumberFormat = "0.00%"
        Workbooks(hw).Sheets(Sheet).Range("Q2:Q3").NumberFormat = "0.00%"
        
        Next Sheet
        
End Sub

Sub delete_all()

    Application.ScreenUpdating = False
    
    For Sheet = 1 To Sheets.Count
        ActiveWorkbook.Sheets(Sheet).Columns("I:Q").Clear
        ActiveWorkbook.Sheets(Sheet).Columns("I:Q").ColumnWidth = 10.38
    Next Sheet
End Sub
