Attribute VB_Name = "Module1"
Sub analysis():

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        Dim tickerName As String
        Dim openValue As Double
        Dim closingValue As Double
        Dim totalVolume As Double
        Dim yearlyChange As Double
        
        totalVolume = 0
        
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        summ_table_row = 2
        
        openValue = 0
        
            For i = 2 To lastRow
                If openValue = 0 Then
                openValue = ws.Cells(i, 3).Value
                
                End If
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    tickerName = ws.Cells(i, 1).Value
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                    closingValue = ws.Cells(i, 6).Value
                    yearlyChange = closingValue - openValue
                    
                    Dim percentChange As Double
                        If openValue = 0 Then
                        percentChange = yearlyChange / 100
                        ElseIf openValue <> 0 Then
                        percentChange = yearlyChange / openValue
                        
                        End If
                        
                        ws.Range("I" & summ_table_row).Value = tickerName
                        ws.Range("J" & summ_table_row).Value = yearlyChange
                        ws.Range("K" & summ_table_row).Value = percentChange
                        
                        If ws.Range("J" & summ_table_row).Value < 0 Then
                            ws.Range("J" & summ_table_row).Interior.ColorIndex = 3
                        ElseIf ws.Range("J" & summ_table_row).Value > 0 Then
                            ws.Range("J" & summ_table_row).Interior.ColorIndex = 4
                        Else
                        End If
                            ws.Range("L" & summ_table_row).Value = totalVolume
                            openValue = ws.Cells(i + 1, 3).Value
                            summ_table_row = summ_table_row + 1
                            totalVolume = 0
                Else
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                End If
            Next i
                ws.Columns("K").NumberFormat = "0.00%"
            
            Dim greatest_percent_inc As String
            Dim greatest_percent_inc_amt As Double
            greatest_percent_inc_amt = ws.Cells(2, 11)
            
            Dim greatest_percent_dec As String
            Dim greatest_percent_dec_amt As Double
            greatest_percent_dec_amt = ws.Cells(2, 11)
            
            Dim greatest_volume As String
            Dim greatest_volume_amt As Double
            greatest_volume_amt = ws.Cells(2, 12)
            
            lastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
            
            For i = 2 To lastRow
                If ws.Cells(i + 1, 11).Value > greatest_percent_inc_amt Then
                    greatest_percent_inc_amt = ws.Cells(i + 1, 11).Value
                    ws.Cells(2, 17).Value = ws.Cells(i + 1, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i + 1, 9).Value
                    
                    End If
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                    
                If ws.Cells(i + 1, 11).Value < greatest_percent_dec_amt Then
                    greatest_percent_inc_amt = ws.Cells(i + 1, 11).Value
                    ws.Cells(3, 17).Value = ws.Cells(i + 1, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i + 1, 9).Value
                    
                    End If
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                
                If ws.Cells(i + 1, 12).Value < greatest_volume_amt Then
                    greatest_percent_inc_amt = ws.Cells(i + 1, 12).Value
                    ws.Cells(4, 17).Value = ws.Cells(i + 1, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i + 1, 9).Value
                    
                    End If
                    ws.Cells(4, 17).NumberFormat = "0"
                Next i
    Next ws
                
    
End Sub
