# Stock Analysis

Analysis of a dozen alternative energy stocks to determine the best investment options, designed to allow for analysis of larger data sets.

## Overview of Project

The scenario for this project is the help a financial advisor named Steve, who wants to advise his clients who are passionate about green energy. The goal is to analyze performance of a dozen green energy stocks over 2017 and 2018 and then optimize the analysis to accomodate larger data sets from the entire stock market.

### Purpose

From the analysis, Steve will be able to analyze the perfromance of a specific stock "DAQO" - already owned by his clients, and the additional investment options by calculating the annualized rate of return. He will be able to run the script to automate his year-on-year analysis of multiple stocks and then re-use it for any stocks in the future. Automating the process will also help reduce the chances of any errors in the repeated analyses.

## Analysis

The data and VBA analysis performed for this project can be found in [Stock-analysis.xlsm](https://github.com/jkenning/stock-analysis/blob/main/VBA_Challenge..xlsm). Below is a short summary of the initial analyses performed:

1. Performed an analysis of the "DAQO" stock in 2018 using loops to calculate daily volume and yearly return

2. Created a flexible macro to perform analysis on all stocks in the data set for any input year, building on the pattern from the DQ analysis

3. Used a macro to customize the fonts, colors, and number formats for the output table to make it easier to read

4. Created buttons to provide Steve with a easy to use interface for running the analyses

5. Modified the all stocks macro to calculate the amount of time it takes to run the script and compile the results

## Results

As a result the stocks data can now be analyzed easily using a button and inputs. In the scenario, Steve now wants to expand the data set to include the entire stock market over the last few years in order to improve his research. As the code works well for a small subset of stocks but may be inefficient or not work well for a larger data set, therefore the aim is to optimize the code through refactoring.

### Original Code

The original code using to analyze all stocks in 2017 and 2018 initializes an array for all the stock tickers then looping through the tickers, with a nested loop to run through the data for each ticker. When running analysis of performance of all stocks in 2017 and 2018 using the original macro, the code ran in **0.5** and **0.51** seconds respectively. 

'find the no of rows to loop over
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop through the tickers
     For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
    
        'loop through the rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
    
        'find total volume for current ticker
            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
        'find starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
            End If
        
        'find ending price for current ticker
        
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
            End If
          
          Next j`

Figure. 1 - Nested loop used in the original code for all stock analysis

### Refactored Code

By refactoring the code it is hoped we can speed up the processing time so that the analysis can handle larger amounts of data in the future. To do this, the code was modified to loop through all the data just one time and collect the same information that would have taken many loops to compile using the previous macro's nested loop. The nested loop is a disadvantage as it takes more time to go through the all the rows for each individual stock. The changes do not add any additional functionality, instead modifying the existing code to take fewer steps and provide a better, more efficient way of accomplishing the same task. Below is the refactored part of the code:

    '1a) Create a ticker Index
    For i = 0 To 11
        
        tickerIndex = tickers(i)
        tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    Next i
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(i) = 0

    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        If Cells(i, 1).Value = tickers(tickerIndex) Then

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = (tickerIndex + 1)
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    
    For i = 0 To 11
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

    Next i`

    Figure 2. - Refactored code removing the nested loop and instead using an index across four different arrarys

In the refactored macro, the single loop uses the tickerIndex variable to reference the ticker arrays for tickerVolumes(12), tickerStartingPrices(12), and tickerEndingPrices(12) - meaning that the tickerIndex is collecting all the information with each new ticker, rather than going through a separate loop to collect information for each ticker every time.

### Refactored Execution Times Comparison

Run time for the refactored code is indeed better for both 2017 and 2018 with around **0.1** and **0.09** seconds respectively (compared to **+/- 0.5** seconds each using the previous code):

![Results for 2017](https://github.com/jkenning/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
Figure. 3 - Performance and run time result for stocks in 2017

![Results for 2018](https://github.com/jkenning/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
Figure. 4 - Performance and run time result for stocks in 2018

### Comparison between 2017 and 2018 Stock Performance

Comparing the results displaying stock performance for 2017 in [Figure 3]() and 2018 in [Figure 4]() it can be seen clearly that 2017 was a much better year than 2018. In 2017 the majority of stocks provided a positive return with the highest values well over 100%, whereas the majority of the same stocks in 2018 provided a negative return up to around -60% at most. EPNH and RUN were the only two stocks that maintained a positive return on investment into 2018, performing exceptionally well over both years, and may represent the best investment choices. By incorporating stock performance data from a wider range of years Steve would be able to better determine long term trends in the performance of the stocks. He can now also determine if there are other good investment options by incorporating data from additional stocks which will be more easily analyzed thanks to the refactored code.

## Summary

### Advantages and Disadvantages of Refactoring Code

**Advantages**
- Improves the design/organization of the code
- Makes the code simpler and easier for others to understand for peer review
- Makes it easier to find and fix bugs
- Easier to re-modify and maintain in the future
- Faster to run (can handle larger data sets)
- Less tied to the order that the code is written (fewer layers)
- Remove redundant code such as duplicates

**Disadvantaes**
- Takes time (and money)
- Existing code may already be fit for purpose and the additional effort may not worth it for only small gains in efficiency
- May introduce new bugs and errors (best to do step-by-step modifications)
- Difficult for large or complex scripts, could sometimes be easier to re-write parts of code from scratch
- Best to do earlier on rather than delayed

### Advantages and Disadvantages Applied to the Original and Refactored VBA Script

The original code was fairly well structured and organized however the main issue was with the lack of efficiency of the nested loop. Using an index to loop through the data instead allowed the nested loop to be removed and is a more efficient alternative to perform the same fucntion. Removing the nested loop makes the code simpler, with less layers dependent on order of execution. If myself, Steve, or another analyst wanted to return and modify the code again in the future it would be much easier to understand and remember what the code does and would require fewer modifications to change its functionality. For the purposes of analyzing just a dozen stocks it could be argued that the amount of time spent refactoring the code and troubleshooting errors that arose as a result was not worth the time saved in faster processing speed, which makes a negligible difference for such a small data set. The key deliverable however is that now Steve can run the code on a far larger number of stocks in the future as a result of the improved processing speed if the size of the data set were to exponentially grow.
