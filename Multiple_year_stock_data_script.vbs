Attribute VB_Name = "Module1"
Sub StockMarketChanges()

' Apply code to each worksheet

For Each ws In Worksheets

    ' Assign Variable Types to All Variables

    Dim LastRow As Long
    
    Dim Summary_Table_Row As Double
    
    Dim RowCount As Double
    
    Dim YearlyChange As Double
    
    Dim YearlyStart As Double
    
    Dim YearlyEnd As Double
    
    Dim StartRow As Double
    
    Dim PercentChange As Double
    
    Dim VolumeTotal As Double
    
    Dim i As Long
    
    ' Assign title names to Summary Values
        
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Yearly Change"
    
    ws.Cells(1, 11).Value = "Percentage"
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    ' Assign last row variable
            
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    ' Assign variable for volume counter - used to sum volume for each stock
    
    VolumeTotal = 0
            
    ' Assign starting row for summary table counter
    
    Summary_Table_Row = 2
            
    ' Assign variable for row counter - used to count number of rows that have progressed for each stock
    
    RowCount = 0
            
    ' Begin for loop within worksheet from row 2 to the last row
            
    For i = 2 To LastRow
                    
        ' If ticker on the next row is different from the current line, proceed
                    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
        ' Add ticker on current line to column 1 on the summary table
                    
        Ticker = ws.Cells(i, 1).Value
                        
        ws.Range("I" & Summary_Table_Row).Value = Ticker
                        
        ' Calculate starting row within range of current stock by subtracting i by the number of rows counted by RowCount
                        
        StartRow = i - RowCount
                            
            ' If stock opening and close values are not equal to 0, proceed
            ' (this ensures we do not get a divide by 0 overflow error create by PLNTs open and close values)
                            
            If ws.Cells(i, 6).Value And ws.Range("C" & StartRow).Value <> 0 Then
            
            ' Assign variables to opening and close values, calculate yearly change, print yearly change in column J
                        
            YearlyEnd = ws.Cells(i, 6).Value
                        
            YearlyStart = ws.Range("C" & StartRow).Value
                        
            YearlyChange = YearlyEnd - YearlyStart
                        
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                            
            ' Calculate percent change, assign percent change to column K and format as a percentage
                            
            PercentChange = YearlyChange / YearlyStart
                                
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
                                
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                
            End If
                                
        ' Calculate Stock Volume
                                
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                                    
        ws.Range("L" & Summary_Table_Row).Value = VolumeTotal
                                    
            ' Assign conditional formatting for yearly change
                                
            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                    
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
            Else
                        
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
            End If
                                    
        '   Add To summary table counter to proceed to next line, reset row counter & volume total
                        
        Summary_Table_Row = Summary_Table_Row + 1
                            
        RowCount = 0
                                        
        VolumeTotal = 0
        
        '  If ticker on current line is not different from ticker on the next line add 1 to row counter
        '  and add value on current line to volume total counter for that stock
                            
        Else
                        
        RowCount = RowCount + 1
                            
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                                    
        End If
                    
    Next i
    
Next ws

End Sub
