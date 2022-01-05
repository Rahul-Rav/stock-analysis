# VBA of Wall Street

## Overview
    Our friend Steve has just graduated with a finance degree and has been asked by his parents to look into DAQO (DQ) stock - a green energy company that they want to invenst in. Steve - with our help is looking into DAQO as well as looking to help his parents diversify their funds. Steve created an Excel sheet containing information on green energy stocks that he is looking for us to analyze. Using VBA, we create code that can automate tasks to help Steve analyze multiple stocks. 
    We are working with stock information from 12 green energy companies within the years of 2017 and 2018, both contained within separate sheets. Using this information, we create a new sheet titled "DQ Analysis" containing information on DAQO (DQ). Through VBA, we wrote code that showed us the Total Daily Volume and Return amount for DAQO in the year 2018 which yielded us less than favorable data. DAQO stock had dropped nearly 63% in 2018, which leads us to write additional code to offer better stock options for Steven's parents using the Excel data provided to us containing information on 11 additional companies. 

## Results
    With the help of our newly written code in VBA, we discover illuminating data on the stocks were given to analyze.
    In 2017, we find that 11 out of the 12 companies were in the green (turned a profit). The profits ranged from 5.55% all the way 199.45%, which so happens to be DAQO.
    Looking back at our previously made sheet and code, we found that in 2018, DAQO stock fell to nearly 63%.
    After running our analysis to display information on 2018, we find that 10 out of the 12 companies were in the red (turned a loss), with DAQO being the biggest loss. So from 2017 to 2018, DAQO went form turning the largest profit (199.45%) to the greatest loss (-62.6).
    With this information, Steve can inform his parents that DAQO may not be the safest option to invest their money in.
    Analyzing these two years further, we discover that only two companies were able to turn a profit in both years.
    ENPH at 129.52% in 2017 and 81.92% in 2018 and RUN at 5.55% in 2017 and 83.95% in 2018.
    While ENPH still turned a profit in 2018, it was 47.6% less than they have made in 2017. RUN on the other hand had small profit in 2017 but increased it by 78.4% in 2018.
    Looking at this data, Steve might consider recommending RUN to his parents to invest in, due to the rate of growth they had made compared to the majority of the other stocks, which 11 out of the 12 had reported losses.


## Summary
    The final portion of this project was the refactoring of our code. Refactoring code is simply the optimization of pre-existing code, which comes with its advantages and potential disadvantages.
    Our original script was successful in running our code to display stock data for any year we choose through the message box. The original code relied on additional complex lines of code that we had to look up, which may not be easy to read for some.
    Through the process of refactoring we were able to successfully make the VBA script run faster. The biggest additions to the refactored code was the use of layered If/Then statements and For loops to optimize the time in which our script runs.
### Advantages of Refactoring    
    Advantages include the potential simplification of pre-existing code to use less memory and run analyses on larger datasets quicker. At the end of our code, we had implemented a  timer to measure code performance. Comparing our initial timer to our refactored code, it was found that the refactored code ran much quicker on our datasets. Beyond the time in which our code runs, refactoring can make the code easier to read and understand. Code can be viwed as a story and the coder as an author - every author has a different style but some may use complex jargon compared to others who may write in ways that are more accessible to less experienced readers.
### Disadvantages of Refactoring    
    Disadvantages of refactoring code are few, but can occur in certain situations. If a code is complex, it can be time consuming to further optimize a line of code. On the other hand, if a code is already simple and efficient and a refactor is asttempted, it is possible for the code to become less efficient, and may be introduced to new bugs.


    
