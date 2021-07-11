# **Module 2 Stock-Analysis Challenge**

## **Overview**

   The purpose of this analysis was that certain stocks wanted to be analysised and checked for potential profitability from a financial advising firm.  We were given two years of stock data for 12 different stock options.  Initially we were asked to look at just one option in particular, DQ, but decided to generalise the macros/code so the client could see the results for not just DQ but also the 11 other options that were in the data set as well.  We created a macro that ran through all the data and collected the total volume of each option that was traded that year.  This macro also collected the starting and ending price of the stock for the year and calculated a return on that option to see if it was a profitable investment over that year.
  
## **Results**

   Upon executing the macro for just the option DQ, we found that we could generalise the macro to better suit the needs of our client.  It would also make use of all the data given to us, not just a section of it.  We also included a timer to get a better sense of how our macro was performing and see if there was any potential for refactoring the code.  The timer did not initialize until the user inputed the year they wished to run the analysis on.  Here are the results of our macro on the years 2017 and 2018, respectively:
  https://github.com/wprich/stock-analysis/blob/main/Resources/Sub%20AllStocksAnalysis.vbs
```
Sub AllStocksAnalysis()
     
   'Timer Variables
    Dim startTime As Single
    Dim endTime As Single   
        
    Worksheets("All Stocks Analysis").Activate

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
```
https://github.com/wprich/stock-analysis/blob/main/Resources/Module_Analysis_2017.png
![Module_Analysis_2017](https://user-images.githubusercontent.com/85487722/125202672-66226a00-e229-11eb-8923-b4883962f3c0.png)

https://github.com/wprich/stock-analysis/blob/main/Resources/Module_Analysis_2018.png
![Module_Analysis_2018](https://user-images.githubusercontent.com/85487722/125202684-73d7ef80-e229-11eb-9f57-f652b04a0fc3.png)

  Seeing as the macro was taking a little longer than desired to run, we decided to refactor the code to see if there was any way to make it run smoother and faster.  This was achieved by createing different arrays to store the results in and a variable tickerIndex and using this as a reference that we would pull our results from said arrays.  The run time had a considerable difference.
  ```
  Sub AllStocksAnalysisRefactored()

    'Initialize timer variables
    Dim startTime As Single
    Dim endTime  As Single

    'Ask user for which year
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Start timer after user input
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    'Create header for results sheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(0 To 11) As Long
    Dim tickerStartingPrices(0 To 11) As Single
    Dim tickerEndingPrices(0 To 11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
        
            '3a) Increase volume for current ticker
                If Cells(j, 1).Value = tickers(tickerIndex) Then
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                End If
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
                If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                End If
        
            '3c) check if the current row is the last row with the selected ticker and increase tickerIndex
                If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                    tickerIndex = tickerIndex + 1
                End If
                
        Next j

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    tickerIndex = i

        Worksheets("All Stocks Analysis").Activate

            Cells(i + 4, 1).Value = tickers(tickerIndex)
            Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
            Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    'Loop through each percentage in Return column and color code if its positive(green) or negative(red)
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i

    'End timer as last loop is ending
    endTime = Timer
    'Display time it took to run said macro in a pop up window
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```


https://github.com/wprich/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png
![VBA_Challenge_2017](https://user-images.githubusercontent.com/85487722/125202794-13957d80-e22a-11eb-8f4a-a0c42fb35201.png)

https://github.com/wprich/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png
![VBA_Challenge_2018](https://user-images.githubusercontent.com/85487722/125202798-1abc8b80-e22a-11eb-8fe9-37414f37153a.png)

## **Results**

  There are several advantages and disadvantages to refactoring code.  First and foremost is that refactoring code can help us to generalise the code and make applicable to multiple scenarios/problems instead of the specific one being worked on.  This can help is reducing man hours needed to solve a problem whose solution closely resembels that of one that has already been completed.  Another advantage is that refactoring code can take new technologies or new coding syntaxs' into account that maybe previously weren't available when the project started.  This includes new patches, keywords, or taking advantage of a new storage option.  
  
   One particular disadvantage to refactoring code is that it can take a considerable amount of time to do, especially when not having the most skilled of labor available.  This could add to project time before completion which can have a negative impact on a companies bottom line.  So the appropriate pros must be weighed against the appropriate cons.  
   
   As for this project and refactoring the code, I think weighing the pros and cons, it was beneficial to take the extra time to refactor the code.  The downside to refactoring the code on this project was that it took about an extra 20-30 minutes to finish.  The use of arrays to store the data in the second set of code is also a best use practice to cut down on the run time of the code.  As the first run of the code will show, it would go through the data set 12 different times and print the results on the analysis page in real time.  As the second set of code runs, it stores each appropriate value in an array for later and then is called back using the tickerIndex variable when populating the analysis page.  This allows it to collect all necessary information on our analysis in one run of the data, not 12.  But now that the code is refactored, it can be applied to a more general data set than the one we were given, and it also takes a considerable less amount of time to run.  And since time is money, this will be greatly appreciated bye the client.
