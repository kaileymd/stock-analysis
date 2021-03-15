# Stock Analysis in VBA

## Project Overview
Steve, a financial advisor, would like to provide data-driven stock trading recommendations to his clients. Using VBA and a subset of available stocks, we've programmatically found the total trading volume and monetary returns of those stocks for a given year. As Steve's career as a financial advisor grows, he'll want to evaluate more stocks and will need to use more efficient programs. 

## Program Efficiency
To illustrate why program efficiency is important, let's dive into the two VBA programs that could give Steve his results: Program A, which used nested for loops for the calculations, and Program B, which only used one for loop to do the same tasks.

### Program A: Nested For Loops

Program A uses nested for loops to comb through the stock data and select the values we need to do later calculations:

![Program A code](https://github.com/kaileymd/stock-analysis/blob/main/Resources/Program%20A%20code.png)

As you can see, the variabe *j* loops through the entire data set to collect values for a given *i*, and it repeats *i* times. In this case, that means *j* needs to loop through the workseet 12 times to collect all of the values it needs to output the stock results.

### Program B: Single For Loop

Program B uses one for loop to collect the same values:

![Program B code](https://github.com/kaileymd/stock-analysis/blob/main/Resources/Program%20B%20code.png)

Using the variables tickerVolumes(12), tickerStartingPrices(12) and tickerEndingPrices(12) as arrays and the static variable tickerIndex, the code is able to start at the beginning of the worksheet and collect the values for each stock as it goes, resulting in only one pass through the stock data. This code utilizes arrays to organize the values and assign it to the correct stock as the tickerIndex increases after each stock value collection is complete.

### Performance Comparison
We input the same data into each program, and the output is the same - so what is there to compare? The time it takes each program to run:

| Program A | Program B |
|:---------:|:---------:|
|![Nested_Loops_2017](https://github.com/kaileymd/stock-analysis/blob/main/Resources/Nested_Loops_2017.png)|![VBA_Challenge_2017](https://github.com/kaileymd/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)|
|![Nested_For_Loops_2018](https://github.com/kaileymd/stock-analysis/blob/main/Resources/Nested_Loops_2018.png)|![VBA_Challenge_2018](https://github.com/kaileymd/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)|

Program B completes its code 1 second faster than Program A, which is ~98% faster.

## Conclusions

Program efficiency is crucial to consider in Steve's case, since the time needed to run the program will increase with the amount of stocks he wants to analyze. Code refactoring is a process that looks to improve a program's code, increasing efficiency and readability. In this program's case, increasing this code's efficiency has greatly improved Steve's ability to quickly analyze even more stocks, advancing the speed and quality of Steve's financial recommendations. While readability isn't as important in Steve's case since he is just running the program, improving readability will help him if he decides to expand the capability of this program in the future.

Improving code through code refactoring may seem like an obviously good thing - it's all about **improving** and making it better, after all! But refactoring may not always be the best decision for a company. A company should consider:
-Is the time investment required to refactor the code worth the efficiency output?
  -For example, if it takes 12 hours to refactor code that will only save 30 minutes a month, is it worth it? What if it saves 30 minutes a day?
-Is this code too outdated and should be replaced? (whether this task is written in another language, irrelevant, etc.)

A program like Program A may be more intuitive for the original coder, or be favored in the event that the specific stocks evaluated change frequently. In Program A, you'd need to update a single array, the index reference in the 'calculations' for loop, and the data output rows - in Program B, you'd need to update the length for 3 arrays, the index reference in the 'calculations' for loop, the index reference in the output for loop, and the data output rows. That makes Program B more difficult to update than Program A. Steve would need to evaluate what program properties are most important to him when deciding which program to adopt.
