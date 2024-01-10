Attribute VB_Name = "Module1"
Sub stock_analysis()

    For Each ws In Worksheets
        ws.Activate
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        row_count = Cells(Rows.Count, "A").End(xlUp).Row
        TotalValue = 0
        openprice = Cells(2, "C").Value
        summarypointer = 2
        
        For i = 2 To row_count
        
            TotalValue = TotalValue + Cells(i, "G").Value
            
            If Cells(i, "A").Value <> Cells(i + 1, "A").Value Then
                closeprice = Cells(i, "F").Value
                yearlychange = closeprice - openprice
                percentchange = yearlychange / openprice * 100
                
                Cells(summarypointer, "I").Value = Cells(i, "A").Value
                Cells(summarypointer, "J").Value = yearlychange
                Cells(summarypointer, "K").Value = "%" & percentchange
                Cells(summarypointer, "L").Value = TotalValue
                
                If yearlychange > 0 Then
                    Cells(summarypointer, "J").Interior.ColorIndex = 4
                ElseIf yearlychange < 0 Then
                    Cells(summarypointer, "J").Interior.ColorIndex = 3
                Else
                    Cells(summarypointer, "J").Interior.ColorIndex = 2
                End If
                
                TotalValue = 0
                openprice = Cells(i + 1, "C").Value
                summarypointer = summarypointer + 1
                
            End If
        Next i
        
    Next ws
End Sub

