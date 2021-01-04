# Module-2-Challenge

## Overview of Project

-In this project I refracted the code to run more efficiently allowing the macro to run faster and to run better at scale. Using the code from the module we took a slightly different approach by using the tickerIndex to better output the data instead of nested If-then statements within the code from the module.

## Results

-As seen in the photos below,the first ones being from the old macro and the next ones being from the refracted code, there is a significant percentage decrease in the amount of time it takes to run. An added bonus is that it is less lines and more straight forward. The second code would be much easier to manipulate if needed to be changed. Accuracy is also key here and we received the same results with both codes. 

![2018 Old Results](Resources/VBA_Challenge_2018_Old.png)![2017 Old Results](Resources/VBA_Challenge_2017_Old.png)
![2018 New Results](Resources/VBA_Challenge_2018.png)![2017 New Results](Resources/VBA_Challenge_2017.png)

### Code

-Here is an example of how the tickerIndex we used within the code to help streamline output instead of using multiple If-then statements. Here we are Increasing volume for the current ticker.
        
```
 '3a) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(J, 8).Value
```
-In the old code this process would be done by looping through the below If-then statement
```
  '5a) Get total volume for current ticker
            If Cells(J, 1).Value = ticker Then
    
                totalVolume = totalVolume + Cells(J, 8).Value
        
            End If
```

## Summary

-Advantages of refactoring this code is the efficiency (speed and memory), the ability to scale, and easier to manipulate if changes need to be made. That being said if this were to be the only application it may not have been worth the time to refactor it to save about .5 secs of time. The new code is also much easier to jump into. It has less moving parts so to speak. All in all I don't see much down to refactoring the code other than time and hassle it may take. 

-Depending on the comfortability with your code and if you have personal preference for Loops and If-then statements then the original code may be the way to go. If the end user has some knowledge of VBA but doesn't know how the tickerIndex works, then this gives the end users something more valuable for their skill level. To me it's more complicated after understanding the function of the tickerIndex that we created. The new code runs smoother with less memory even though it is hard to tell the difference now it will come in hand in the future if we decide to expand the amount of stocks we look at and the amount of data.



