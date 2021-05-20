Sub Summary()

For Each ws In Worksheets

    Dim ticker As String
    
        Dim ticker_total As Double
            
            ticker_total = 0
    
    Dim summary_table_row As Integer
    
        summary_table_row = 2
        
    Dim open_value As Double
        
        open_value = ws.Range("C2").Value
    
    Dim close_value As Double
    
    Dim price_change As Double
    
    Dim percent_change As Double
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    ws.Range("I1").Value = "Ticker"
    
    ws.Range("J1").Value = "Yearly Change"
    
    ws.Range("K1").Value = "Percent Change"
    
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    ws.Range("O2").Value = "Greatest % Increase"
    
    ws.Range("O3").Value = "Greatest % Decrease"
    
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("P1").Value = "Ticker"
    
    ws.Range("Q1").Value = "Value"


    For i = 2 To lastrow
    
        If Not IsEmpty(ws.Cells(i, 3)) Then
    
            If ws.Cells(i, 3).Value = 0 Then
    
                ws.Rows(i).EntireRow.Delete
        
                i = i - 1
            
            End If
        
        End If
                
    Next i
    
    

    For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
            ws.Range("I" & summary_table_row).Value = ticker
            
            ws.Range("L" & summary_table_row).Value = ticker_total
            
            ticker_total = 0
            
        Else
        
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
        End If
        
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            

                close_value = ws.Cells(i, 6).Value
                
                price_change = close_value - open_value
                
                percent_change = (close_value / open_value) - 1
                
                open_value = ws.Cells(i + 1, 3).Value
                
                ws.Range("J" & summary_table_row).Value = price_change
                
                
                    If ws.Range("J" & summary_table_row).Value > 0 Then
                    
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                        
                    ElseIf ws.Range("J" & summary_table_row).Value < 0 Then
                    
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                        
                    End If
                    
                
                ws.Range("K" & summary_table_row).Value = percent_change
                
                    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                
                summary_table_row = summary_table_row + 1
            
           
            
            
        End If
        
         
    Next i
    
    
'Bonus_______________________________________________________________________________

    
    last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim top_percent_change As Double
    
        top_percent_change = 0
        
    Dim top_percent_change_ticker As String
    
    Dim low_percent_change As Double
    
        low_percent_change = 100
        
    Dim low_percent_change_ticker As String
    
    Dim top_volume As Double
        
        top_volume = 0
        
    Dim top_volume_ticker As String
    
    
    
    For i = 2 To last_row2
    
        If ws.Cells(i, 11).Value > top_percent_change Then
        
            top_percent_change = ws.Cells(i, 11).Value
            
            top_percent_change_ticker = ws.Cells(i, 9).Value
         
        End If
        
        If ws.Cells(i, 11).Value < low_percent_change Then
        
            low_percent_change = ws.Cells(i, 11).Value
            
            low_percent_change_ticker = ws.Cells(i, 9)
            
        End If
        
        If ws.Cells(i, 12).Value > top_volume Then
            
            top_volume = ws.Cells(i, 12).Value
            
            top_volume_ticker = ws.Cells(i, 9).Value
            
        End If
         
        
    Next i
    
    ws.Range("Q2").Value = top_percent_change
        
        ws.Range("Q2").NumberFormat = "0.00%"
        
    ws.Range("P2").Value = top_percent_change_ticker
    
    ws.Range("Q3").Value = low_percent_change
    
        ws.Range("Q3").NumberFormat = "0.00%"
        
    ws.Range("P3").Value = low_percent_change_ticker
    
    ws.Range("Q4").Value = top_volume
    
    ws.Range("P4").Value = top_volume_ticker
    

    
Next ws

End Sub