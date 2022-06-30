Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
    
    ' Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
   
    ' Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    ' Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    ' Activate data worksheet
    Worksheets(yearValue).Activate
    
    ' Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Create a ticker Index
    tickerIndex = 0
    
    'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    ' Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
   ' Loop through rows in the data
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
   
        ' Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        ' Check if the current row is the first row with the selected tickerIndex
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
            
        End If
            
        ' Check if the current row is the last row with the selected ticker
        ' If the next row's ticker doesn't match, increase the tickerIndex
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            ' Assign value to the tickerEndingPrice(tickerIndex)
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            
            ' Increase the tickerIndex
            tickerIndex = tickerIndex + 1
            
        End If
        
    Next i
            
    ' Loop through your arrays to output the Ticker, Total Daily Volume, and Return
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        
        'Record the tickers information in our summary table
        'Record the ticker
        Cells(4 + i, 1).Value = tickers(i)
        
        'Record the tickerVolumes
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'Record the return
        Cells(4 + i, 3).Value = (tickerEndingPrice(i) / tickerStartingPrice(i)) - 1
        
    Next i
     
    ' Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
        For i = dataRowStart To dataRowEnd
            If Cells(i, 3) > 0 Then
            
                'Color the cell green
                Cells(i, 3).Interior.Color = vbGreen
            ElseIf Cells(i, 3) < 0 Then
            
                'Color the cell red
                Cells(i, 3).Interior.Color = vbRed
            Else
            
                'Clear the cell color
                Cells(i, 3).Interior.Color = xlNone
            End If
        Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub
