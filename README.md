# ALL STOCK ANALYSIS IN EXCEL WITH VBA

## OVERVIEW OF PROJECT

## PURPOSE
The aim of this project was to refactor a Microsoft Excel VBA code to collect certain stocks information in the year 2017 and 2018 and analyze whether or not the stocks are worth investing. Both of this VBA code was originally coded in a similar format, however, the goal for this refactor was to make it more efficient and lessen run time execution to help our client to visualize outcome quicker.

## DATASET OVERVIEW
The information that is retreived incorporates two results dependent on year 2017 and 2018 with stock data on 12 distinct stocks. The stock data contains a ticker esteem, the date the stock was given, the opening, closing and changed closing price, the most noteworthy and least price, and the volume of the stock. The objective is to recover the ticker, the all out day by day volume, and the profit from each stock.

## OUTPUT RESULTS

## ANALYSIS 
Prior to refactoring the code, I started by duplicating the code that was expected to create the ticker array, and to activate the specified worksheet. The means were then rattled off to set the construction for the refactoring. The snippet of the code remarked with brief clarification can be found underneath:

	1a) Started off by initializing the ticker Index to start from 0
	tickerIndex = 0

	1b) Here, creating the required output arrays
	Dim tickerVolumes(12) As Long
	Dim tickerStartingPrices(12) As Single
	Dim tickerEndingPrices(12) As Single

	2a) Create a for loop to initialize the tickerVolumes to zero.
	For i = 0 To 11
    	tickerVolumes(i) = 0
    	tickerStartingPrices(i) = 0
    	tickerEndingPrices(i) = 0
	Next i

	2b) Then looped over all the rows in the spreadsheet.
	For i = 2 To RowCount

    	3a) Increase volume for current ticker
    	tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    	3b) Check if the current row is the first row with the selected tickerIndex.
    	If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    	End If
    
    	3c) check if the current row is the last row with the selected ticker
     	If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     	End If

        3d) Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

	Next i

	4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
	For i = 0 To 11
   	    Worksheets("All Stocks Analysis").Activate
   	    Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
        Next i

## CONCLUSION

## ADVANTAGES AND DISADVANTAGES OF REFACTORING CODE
Refactoring helps make our code cleaner and more organized.
A couple of benefits of ensuring clean code incorporate plan and programming improvement, troubleshooting, and quicker programming. It might likewise profit different clients who see our tasks since it gets simpler to peruse, as it is more brief and clear.
In any case, we don't generally have the privilege to refactor our code because of detriments. These detriments may go from having applications that are too enormous to not having the legitimate use and test cases for the current codes, which may hinder refactoring.

## ADVANTAGES AND DISADVANTAGES OF THE ORIGINAL AND REFACTORED VBA SCRIPT
The greatest advantage that happened because of the refactoring was a decline in full scale run time. The original analysis required roughly about 0.17 to 0.23 seconds to run, though our new refactored script just required approximately around 0.16 seconds to run. All the screen captures that demonstrate the run time for our new refactored analysis can be found down below:

OUTPUT BEFORE REFACTORING
![stock2017Before](https://user-images.githubusercontent.com/86158802/125212191-b0333c00-e279-11eb-86dc-22f8112b43fc.PNG)
![Stock2018Before](https://user-images.githubusercontent.com/86158802/125212208-c214df00-e279-11eb-8d1f-b6436af51c4f.PNG)

OUTPUT AFTER REFACTORING
![VBA_Challenge_2017](https://user-images.githubusercontent.com/86158802/125212220-d5c04580-e279-11eb-9a61-81c6ee69b82c.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/86158802/125212221-d8bb3600-e279-11eb-977f-799e6ee1164c.PNG)



