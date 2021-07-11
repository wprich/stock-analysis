Sub AllStocksAnalysis()
        
    Worksheets("All Stocks Analysis").Activate
    
   'Timer Variables
    Dim startTime As Single
    Dim endTime As Single
        'Ask for user input
    yearValue = InputBox("What year would you like to run the analysis on?")
        
        'Timer Start
            startTime = Timer
            
        'Set Header for Results sheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
        'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
        'Setting up array for tickers
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
        
        'For loop to go through each ticker
    For i = 0 To 11
    
        'Ensures 2018 sheet is selected
            Worksheets(yearValue).Activate
    
        'Variables declarations
            totalVolume = 0
            'rowEnd found on https://stackoverflow.com/questions/18088729/row-count-where-data-exists/18089140
            rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
            rowStart = 2
            Dim startPrice As Double
            Dim endPrice As Double
            
        'Assign the appropriate ticker
            ticker = tickers(i)
        
            'Check for ticks to add to totalVolume
                For j = 2 To rowEnd
               
            'Increases totalVolume for each tick counted
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
            
            'Set Starting Price
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                    startPrice = Cells(j, 6).Value
                End If
            
            'Set Ending Price
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                    endPrice = Cells(j, 6).Value
                End If
            
                Next j
        'Output results
           Worksheets("All Stocks Analysis").Activate
           
           Cells(i + 4, 1).Value = ticker
           Cells(i + 4, 2).Value = totalVolume
           Cells(i + 4, 3).Value = endPrice / startPrice - 1
    Next i
    
        'Formatting
            Worksheets("All Stocks Analysis").Activate
            Range("A3:C3").Font.Bold = True
            Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range("B4:B15").NumberFormat = "#,##0"
            Range("C4:C15").NumberFormat = "0.0%"
            Columns("B").AutoFit
            For i = 4 To 15
    
                If Cells(i, 3).Value > 0 Then
                    Cells(i, 3).Interior.Color = vbGreen
        
                ElseIf Cells(i, 3).Value < 0 Then
                    Cells(i, 3).Interior.Color = vbRed
            
                Else
                    Cells(i, 3).Interior.Color = xlNone
            
                End If
        
            Next i
            
    'Timer Stop
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
End Sub
