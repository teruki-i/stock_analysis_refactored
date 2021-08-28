# Testing the Efficiency of Refactored VBA Script

## Overview of Project

### Purpose

This project sought to refactor a VBA script and make it run more efficiently. For this project, I refactored a script used for analyzing stock performance in a given year.

### Results

With the original VBA script, the analysis took 0.97 seconds to completely analyze 2017 stock data. However, with the refactored code, the same analysis took approximately 0.82 seconds to complete.

![original_2017](https://github.com/teruki-i/stock_analysis_refactored/blob/91e1b41d46a9a7208b7c2762d01e271a02b32440/resources/Original_2017.png)

The image above is of the 2017 output with the time measurement for the original script.

![refactored_2017](https://github.com/teruki-i/stock_analysis_refactored/blob/91e1b41d46a9a7208b7c2762d01e271a02b32440/resources/VBA_Challenge_2017.png)

The image above is of the 2017 output with the time measurement for the refactored script.

The results were similar with the 2018 stock data analysis. With the original script, the analysis took approximately 0.95 seconds to complete, but took 0.83 seconds to complete with the refactored script.

![original_2018](https://github.com/teruki-i/stock_analysis_refactored/blob/main/resources/Original_2018.png)

The image above is of the 2018 output with the time measurement for the original script.

![refactored_2018](https://github.com/teruki-i/stock_analysis_refactored/blob/main/resources/VBA_Challenge_2018.png)

The image above is of the 2018 output with the time measurement for the refactored script.

Based on these results, it is clear that the refactored script runs faster, though only by about 0.12 to 0.15 seconds. However, for this analysis, both the 2017 and 2018 data sets had 3012 rows of data and only focused on 12 different stocks. So while the difference in efficiency might seem insignificant for such a limited data set, a more extensive analysis, for example one that contained the whole stock market or one that possibly examined multiple years' worth of data at once rather than just one year, the difference could be more dramatic.

This difference in performance might be because of the difference in how the results were outputted by each script.

    For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

        'activate data worksheet
        Sheets(yearValue).Activate
        For j = rowStart To rowEnd

            'increase totalVolume for each ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If

            'set starting price for ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If

            'set ending price for ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If

        Next j

        'activate output worksheet
        Worksheets("All Stock Analysis").Activate

        'label row for each stock ticker
        Cells(4 + i, 1).Value = ticker

        'output results
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i

In this portion of the original script, the results for each stock were calculated and then added into the output worksheet. As part of the nested for loop though, the script had to constantly switch between the data and output worksheets. Right before the inner loop initiated, the data worksheet had to be activated and right after, the output worksheet had to be activated.

The refactored script, however, doesn't require this constant switch between worksheets because it used arrays

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

These arrays were used to store the volumes, starting prices, and ending prices of the stocks. This was the key to making the nested for loop run more efficiently.

        For j = 2 To RowCount

        '3a) Increase volume for current ticker
            If Cells(j, 1).Value = ticker Then

                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

            End If

        '3b) Check if the current row is the first row with the selected stock ticker. If yes, sets starting price for selected ticker.

            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

            End If


        '3c) check if the current row is the last row with the selected stock ticker.

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'if yes, sets ending price for given ticker
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

                '3d) if yes, increases tickerIndex
                tickerIndex = tickerIndex + 1

            End If

        Next j

    Next i


In this portion of the refactored script, the nested for loop was modified in such a way that there isn't a constant switch between the two worksheets. In fact, no worksheet is activated within the loop

Because the original script didn't use arrays, it was necessary to output the results before the values were overwritten with results for another stock by another iteration of the inner loop. This was why the original script required constantly switching between the two worksheets. However, because the arrays in the refactored script stored the volumes and prices for all the stocks, there was no need to immediately output the results for each stock. The output was done in a separate process from this nested for loop

### Summary

This project demonstrates that refactoring can be beneficial in that it can make a more efficient script without having to write one from scratch. Since refactoring entails working off an already working script, the basic outline and structure are already there. There's also no need to start fresh and come up with the fundamental logic behind the script either because of this. While refactoring, one would just need to find specific pieces of the script that can be improved and run more efficiently.

As discussed above, refactoring made the script more efficient by using arrays. By storing values in arrays, it was possible to modify the nested for loop so that there was no need to constantly switch between worksheets.

However, the disadvantage is that the script can become more complicated and therefore more prone to error. While the point of refactoring is to create a script that runs more efficiently, it might require replacing segments of a script that, while less efficient, are much easier to follow.

For example, by using arrays, in the refactored script, it became necessary to specify an index to make sure that the correct piece of data gets stored within each slot the arrays. Otherwise, the array could be overwritten as a variable with just one value. Furthermore, it also became necessary for the index for the arrays to increase when necessary. In order for this to happen, the index had to be initialized as an integer equal to 0 and be part of a conditional statement in a for loop in order to increase. Missing any of these steps would have made the nested for loops in the script not function correctly whereas the original script kept it much simpler with no need of indexes.
