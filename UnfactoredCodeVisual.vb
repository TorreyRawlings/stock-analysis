Worksheets("All Stock Analysis").Activate 'telling the code to use the 'All Stock Analysis' worksheet
    
        Range("A1").Value = "All Stocks (" + YearValue + ")" 'inputs 'All Stocks (*Year inputted above)' into cell A1
        Range("A3").Value = "Ticker"
        Range("B3").Value = "Total Daily Volume"
        Range("C3").Value = "Return"
        
'2.Initialize an array of all tickers.
Dim tickers(12) As String 'creating a string each ticker is a stock from the worksheet
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

'3.Prepare for the analysis of tickers.
    'Initialize variables for the starting price and ending price.
    Dim startingprice As Single 'stating what kind of variable. this is telling it to use a whole number
    Dim endingprice As Single
    'Activate the data worksheet
Sheets(YearValue).Activate 'tellingit which worksheet to use for outer loop *Year inputted above
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row 'telling code to count the rows with data in it
'4.Loop through the tickers.
    For i = 0 To 11 'saying to count everything from 0 to 11 if using 'i'
        ticker = tickers(i) 'stating when 'ticker' is used to reference the 'tickers' array/list we made above and to reference 'i' which is 0 - 11
        totalVolume = 0 'creating the variable total volume and setting it at 0
'5.Loop through rows in the data.
Sheets(YearValue).Activate 'telling which worksheet to use for the inner loop *Year inputted above
        For j = 2 To RowCount 'creating a new veriable and a inner loop, j is going to count all the rows starting at 2 to all the rows with data
    
    'Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then 'condition stating that if cells in column 'A' are one of the tickers about then do the following ->
            totalVolume = totalVolume + Cells(j, 8).Value 'telling us that totalvolume will now be '0' (as stated above) PLUS column 'H's value for each row that has data (all added together)
        End If 'ends the condition and statement
    'Find the starting price for the current ticker.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then 'condition stating that if the above cell is not one of the tickers but the current cell is then do the following ->
            startingprice = Cells(j, 6).Value 'stating the startingprice will pull what is in column 'F' as the starting price (this is just the 1 cells data, this doesnt add it together)
        End If 'ends the condition and statement
    'Find the ending price for the current ticker.
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then 'condition stating that if the below cell is not a ticker but the current cell is then do the following ->
            endingprice = Cells(j, 6).Value 'stating the endingprice will pull what is in column 'F' as the ending price (this is just the 1 cells data, this does not add it together)
        End If 'ends the condition and statement
        Next j 'closes inner loop
'6.Output the data for the current ticker.
    Worksheets("All Stock Analysis").Activate 'activates the sheet for the outerloop
    Cells(4 + i, 1).Value = ticker 'inputs the ticker info 4 rows down and then to list them all in column A
    Cells(4 + i, 2).Value = totalVolume 'inputs the total volume for each ticker starting 4 rows down, in column B
    Cells(4 + i, 3).Value = endingprice / startingprice - 1 'inputs the ending price divided by the starting price subtract 1 for each ticker starting 4 rows down, in column C
    
Next i 'closes outer loop
    
    endTime = Timer 'end the timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year" & (YearValue) 'message box to input how long the code took the run
End Sub
