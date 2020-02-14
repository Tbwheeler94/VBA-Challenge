Attribute VB_Name = "Module1"
Sub StockMarketChanges()

For Each ws In Worksheets

    Dim LastRow As Long

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Yearly Change"
    
    ws.Cells(1, 11).Value = "Percentage"
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
        Ticker = ws.Range("A2:A" & LastRow)
    
        ws.Range("I2:I" & LastRow) = Ticker
            
            Dim i As Long
            
            For i = 2 To LastRow
                    
                    ws.Range("J" & i).Value = ws.Range("F" & i).Value - ws.Range("C" & i).Value
                    
                    If ws.Range("J" & i).Value <> 0 Then
                    
                        ws.Range("K" & i).Value = (ws.Range("J" & i).Value / ws.Range("C" & i).Value) * 100
                    
                    End If
                    
                     If ws.Range("J" & i).Value > 0 Then
                    
                        ws.Range("J" & i).Interior.ColorIndex = 4
                        
                    Else
                        
                        ws.Range("J" & i).Interior.ColorIndex = 3
                        
                    End If
                    
            Next i
            
                
            TotalStockVolume = ws.Range("G2:G" & LastRow)
    
            ws.Range("L2:L" & LastRow) = TotalStockVolume
    
Next ws

End Sub

