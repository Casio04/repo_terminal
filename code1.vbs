Sub credit_card()

    Dim brand_name As String
    Dim brand_total As Double
    Dim summary_table_row As Integer
    
    summary_table_row = 2
    
    For i = 2 To 101
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            brand_name = Cells(i, 1).Value
            brand_total = brand_total + Cells(i, 3).Value
            Cells(summary_table_row, 7).Value = brand_name
            Cells(summary_table_row, 8).Value = brand_total
            summary_table_row = summary_table_row + 1
            brand_total = 0
        Else
            brand_total = brand_total + Cells(i, 3).Value
            
        End If
        
    
    Next i
    
End Sub
