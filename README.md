# Stock Analysis Using Excel VBA

## Overview of Project

### Background

The dataset shows a set number of tickers, their opening and closing prices, their highest and lowest prices, and finally their volume. They were all taken from the years of 2017 and 2018. This dataset has been sorted in advance but needs the running time to be reduced.

### Purpose

The purpose is to write a VBA script that shows the total volume of stocks per ticker over the year including their return. In addition, the vba script needed to run a refractor code so that the code loops through the data once, reducing the running time of the script.

## Results

### Analysis

Using the previously written code, steps were made to make the script more effecient reducing the run time. Copying over the original code, edits were made to the index and removing nested loops. Below is the code used. 

'1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
              
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
            
            End If
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i

Once the code was ran it properly displayed the tickers, total volume, as well as their return in the years of 2017 and 2018 respectively. Because the code was refactored and became more effecient the running time for the code was greatly reduced. 

![Chart](https://raw.githubusercontent.com/Drakeblaze10/stock-analysis/main/Resources/VBA_Challenge_2017.PNG)
![Chart](https://raw.githubusercontent.com/Drakeblaze10/stock-analysis/main/Resources/VBA_Challenge_2018.PNG)

## Summary

### Pros and Cons of Using Refactoring Code
**Pros:** Nested loops are removed, allowing the code to loop once. This greatly reduces the run time. Previously the code would take aproximately 0.8 seconds to the now 0.1 seconds. Lines of code are more readable with less repeating code creating cleaner script.

**Cons:** Creating more efficient code allows the introduction of more errors, requiring more time for debugging. Another potential advantage may fall when more tickers are added. Creating, for example, a list of 1000 tickers may be arduous, and there may need an easer way to write it all down. 

The benefit that refactoring Code has over nonrefactored is the speed and efficiency of the script. Some may enjoy seeing more individual lines of code written out, breaking down each step. However, with refractoring lines of code can be written more concisely.
