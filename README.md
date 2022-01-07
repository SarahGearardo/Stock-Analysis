# Stock-Analysis

## Project Overview</br>
The original purpose of this project was to analyze the DQ stock for our friend Steve's parents. They were interested in investing in this particular stock and wanted to make sure that they would get the most out of their money. When we analized the DQ stock, it was found that the return rate was poor for their overall volume. Due to that, we decided to focus on reviewing the rate of return for 2017 and 2018 to see which stocks have the best return rate and hopefully give that information to Steve's parents to make a better decision on which stock to pick. After we were done coding the spreadsheet and making a button for ease of use, the data was refractoroed to make it run a little faster.</br>

## Results</br>
### 2017</br>
In our analysis of 2017 data, we find that all but one of the stocks had positive returns that year. The only stock that under-performed was TERP. The highest rate of return that year was DQ, which is what Steve's parents originally wanted to invest in. The return for DQ that year was 199.4%.

![2017 Return Data](https://user-images.githubusercontent.com/94804527/148479346-656529a5-956f-4628-8177-eba9db32f1e8.png)

The runtime of the original 2017 report was .86 seconds:

![2017 Run Time - Not Refracted](https://user-images.githubusercontent.com/94804527/148479396-f1d85d61-2c71-4f05-a9d8-6b8776536fb9.png)

The runtime of the refactored 2017 code is 99 seconds:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/94804527/148485410-bce88f80-af72-402b-ac02-ccb03badf062.png)

### 2018</br>
In our analysis of the 2018 data, we find that most of the stocks had negative return rates, except for two. ENPH had a return of 81.9% and RUN had a return of 84.0%. When comparing the two stocks, ENPH seems like the safest bet as they both had positive return rates in both 2017 and 2018.

![2018 Return Data](https://user-images.githubusercontent.com/94804527/148482449-ba73d98a-645d-49f3-bc25-bcff1dcb0391.png)

The runtime of the original 2018 report was .84 seconds.

![2018 Run Time - Not Refracted](https://user-images.githubusercontent.com/94804527/148485137-ff5ab666-4180-49ba-9a68-8942761989a5.png)

The runtime of the refactored 2018 code is 1.03 seconds:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/94804527/148485502-4d35981e-83cb-46c9-a47f-4ba3cd4822a7.png)

### Code</br>

The code I used to refract the code we were given is displayed below:

 '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = "0"
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
        
    ''2b) Loop over all the rows in the spreadsheet
    Worksheets(yearvalue).Activate
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = "tickerIndex" And Cells(i - 1, 1).Value <> "tickerIndex" Then
        tickerStartingPrices(12) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = "tickerIndex" And Cells(i + 1, 1).Value <> "tickerIndex" Then
        tickerEndingPrices(12) = Cells(i, 6).Value
        End If

            '3d Increase the tickerIndex.
        If Cells(i, 1).Value = "tickerIndex" And Cells(i - 1, 1).Value <> "tickerIndex" Then
        tickerIndex = tickerIndex + 1
        End If
    
    Next i
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
       
    Next i

## Summary</br>
### Advantages and Disadvantages in Refractoring</br>
I believe the advantages of refractoring are easy to sell, which is that refratoring your code can help it run faster, although I did not find that was true with this assignment. This might be useful if you are working with a lot of data at once and I can see how it could easily shave minutes off of the time it takes to run your script to get the data you need. Another advantage as well as disadvantage to refratoring your own code is that you are familiar with it, so sometimes the changes you need to make aren't always going to be obvious to you if you have been working on it for a long time, however, the same can be said for knowing your own code inside and out. If I was to refractor someone elses code without building it myself first, I can see how it could take longer because I would need to get familiar with how they "speak".</br>

### Advantages and Disadvantages on the Refractored VBA Script</br>
The biggest advantage I found with working with the refractored script is that there was less text to work with. I felt like the original script we built to get the analysis for the DQ stocks, which eventually led us to analyzing all stocks was getting a bit combourome. One disadvantage I found with refractoring this VBA script was having to have two modules open for differet sets of code. If you aren't paying attention, I can understand how it could be confusing on which module you are working with. There were times I thought I was running code with my refractored modle, but it turned out to be the original code because I was not paying close enough attention. This could lead to mistakes if people don't catch themselves. 
