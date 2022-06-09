# Stock_Analysis_Challenge

## Refactoring VBA code that analyzes and formats stock data

### Project Overview
My Client (Steve) is helping his parents invest in the green energy sector by analysing the performance of multiple stocks. To expidite the process, I wrote a VBA script that allows Steve to select which year's data he would like to analyse, and after selecting, the script extracts the necessary data, calculates the Total Daily Volume and Total Return for each stock, and then formats these results such that they can be easily interpreted. The origional VBA script worked for the stocks that Steve initially provided, however, he now intends to add more stock data to the Excel and wants the script refactored (_minimizing the time required to perform future analysis_). The following reviews how I refactored this script, and the subsequent results of refactoring.

### Results
As intended, the refactored script allows the user to select which year of stock data to analyse `yearValue = InputBox("What year would you like to run the analysis on?")`, sifts through and calculates the Total Daily Volume and Total Return of each ticker from that year's stock data, formating the results on a new worksheet such that the user can then make easily informed decisions. Both the origional script and the refactored script produce the same results, however, the refactored script takes significantly less time to execute. As demonstrated here, the origional script was inefficient in that it ran through the dataset twelve times as instructed by the following nested 'for' loop. 

````
    'Loop for switching to the next ticker
    For i = 0 To 11
        
        'set TotalVolume to zero for next ticker
        ticker = tickers(i)
        TotalVolume = 0
        Worksheets(yearValue).Activate
        
        'Loop through all rows in the speadsheet, extracting data for current ticker
        For j = 2 To RowCount
            
            'Increase TotalVolume for current ticker
            If Cells(j, 1).Value = ticker Then
                TotalVolume = TotalVolume + Cells(j, 8).Value
            End If
            
            'Assign startingPrice for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'Assign endingPrice for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
        
        'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = TotalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
````

The problem with the above script is in its method (_or lack thereof_) for storing data. Because of this, the script runs through the entire worksheet _**Twelve**_ times, collating the data for one ticker at a time and outputting this information before repeating the entire process for the next ticker. Thus, the first step in refactoring this code was to establish three new arrays for each ticker's volume. In addition to the initial array of tickers, 
````
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
````
I created an array for storing/calculating the volume `tickerVolumes(12)`, starting price `tickerStartingPrices(12)`, and ending price `tickerEndingPrices(12)` of each ticker. Additionally, I established a new variable ``tickerIndex`` that is used to inform the script which data should correspond with which ticker. After creating these new arrays, (_and after being sure to initialize the ticker index to zero, and using a 'for' loop to initialize all ticker volumes to zero_)
````
    'For loop initializing all tickerVolumes to zero
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
````
now, the script only needs to run through the dataset once, as it is capable of storing all outputs within these newly established arrays. As outlined in the code below, the tickerIndex helps the script to understand which data should correspond with different tickers, as it is set to increase once the script recognizes that the ticker within the row below no longer matches the ticker within the current row. Thus, the tickerIndex effectively starts at zero, and increases each time it encounters a new ticker, informing the other arrays that this new data should be stored separately. 

````
    Next i
            
    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        'Increase tickerVolume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
            'Check if the current row is the first row with the selected tickerIndex
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                
                'If first row, Set tickerStartingPrice
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
            End If
        
            'Check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'If the next row’s ticker doesn’t match, set tickerEndingPrice
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                'if next row's ticker doesn't match, Increase the tickerIndex
                tickerIndex = tickerIndex + 1
       
            End If
    
    Next i
````
The beauty of this change to the code, is that once this loop has been completed, all the data necessary for the anlysis worksheet are already effectively stored and ready for output. With the following code, we area able to populate our output table. 
````
    'Loop through arrays to output data for the Ticker, Total Daily Volume, and Return
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
````
Leveraging the new arrays, which I mentioned above, the script populates the stored data within the table. And then, (_as did the origional script_) the script formats the results such that it can be easily interpreted by the user, and finally, produces a message box that informs the user how long the script took to run.

_Run time for the Origional Script analysing stocks from 2017 and 2018, respectively._
(_images below_)

![VBA_Challenge_UNrefactored_2017](https://user-images.githubusercontent.com/106599446/172875527-c10d6b85-0582-42e9-936d-3eaea51befd3.png)
![VBA_Challenge_UNrefactored_2018](https://user-images.githubusercontent.com/106599446/172875654-20db34e8-6308-46d4-96a7-866fbfd3b399.png)


_Refactored Script analysing stocks from 2017 and 2018, respectively._
(_images below_)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106599446/172875548-e62ae66d-ef3d-4c80-8a80-98cc424ba1c2.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106599446/172875604-745fa5a6-f107-4568-94fd-68a074ba4753.png)

As shown above, the refactored script took significantly less time to run. As mentioned above, the refactored code was much more efficient in that it needed to run through the dataset only once, whereas the origional script ran through the dataset twelve times. Despite taking less time to run, the refactored script produces the exact same results.

Results for 2017 stocks

![VBA_Challenge_Results_2017](https://user-images.githubusercontent.com/106599446/172875164-b56fc264-3277-4a66-aa72-52d5d02fda2c.png)

As seen in the above table,

Results for 2018 stocks

![VBA_Challenge_Results_2018](https://user-images.githubusercontent.com/106599446/172875224-94355335-d377-4ce9-b0e7-80a059b37f18.png)






### Summary
- Advantages / Disadvantages
- Applying pros & cons to refactoring origional VBA script
