# stock-analysis
stock-analysis for module 2: VBA Challenge

## VBA Challenge Overview
The purpose of this challenge was to refactor the code of a stock-analysis macro, so that the runtime is faster than before. The workbook is filled with two sheets of yearly return data from twelve stocks. The macro is meant to sort through daily returns of a stock throughout the year and return the total daily volume along with a return percentage for said stock. This is a lot of data to go through, so it is imperative to have an efficient macro.

## Results
After refactoring the code, the purpose of this challenge was a success. With the new code, the macro runs about 10 times faster than before. Originally it would take approximately 0.5 seconds to run. Now it takes approximately 0.05 seconds to run. With this amount of data it may not seem that is a big difference, but with a larger data set it would be very noticeable. Overall the new code performs much better.

<img width="793" alt="2017runtime" src="https://user-images.githubusercontent.com/104074135/178914235-04275cc5-6339-420e-9b87-aa0224db859f.png">
<img width="790" alt="2017Oldruntime" src="https://user-images.githubusercontent.com/104074135/178914575-c9b40601-9a47-44c8-b6b8-1ab939d7fa55.png">

###### In Depth
The theory behind refactoring was to take all of the code from the data sheet instead of going back and forth to receive data. This was possible by adding output arrays to run through the data we need rather than go back to the previous sheet to pull the next ticker. 

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
``` 

Previously, the macro ran a nested for loop. By adding these new output arrays I was able to remove the nested for loop. This means instead of looping through the data twelve times, I only had to loop through it once and went through each array ticker in the same loop.

```
 For j = 2 To RowCount
      
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                
            End If
                  
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            End If
            
            
        'End If
    
    Next j
```

## Summary
The key advantage here is that the macro runs more efficiently. It will be less demanding for the application and computer to run. Also, a faster runtime benefits the user. The disadvantage of refactoring code is that the original code may not always be something a person is familiar with. Thankfully in this module I was the one to go through the original code, and I had comments to keep track of things. Comments are another advantage while refactoring code because they give insight on the intention of the written code and what it is meant to do. Overall, the new macro was a success by decreasing runtime while preserving all intended functions.
