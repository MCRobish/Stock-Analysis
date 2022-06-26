# Stock-Analysis
## Overview of the Project
Steve's parents would like to invest their money in renewable energy, but Steve has some concerns about the risk of low returns for the company they have selected. Steve would like to have an efficient macro that is able to analyze the results of many stock prices over multiple years. This would allow him to choose some stocks that have lower risk but still meet their parents interests in green energy. The inital code used looped through the dataset multiple times which is too slow for a large data set. The goal is to reduce the time to execute this code while maintaining accuracy. The initial code is stored in this repository under **green_stocks.xlsm.** The refactored macro with faster results is stored in the _challenge_ folder in this repository as **refactoring of greenstocks challenge.xlsm.** 

## Results
**Stock Performance**

The stock returns and total daily volume are shown below for the two years in the dataset. 
<p align="center" width="100%">
    <img width="33%" src=https://user-images.githubusercontent.com/105991478/175822019-a7637763-b0b0-4139-ad8b-1598ff3d0d40.png>  <img width="33%" src= https://user-images.githubusercontent.com/105991478/175822020-e5334d74-34d1-48d3-a373-82120080861d.png>
</p>

Overall, the returns for 2017 were significantly better for the green energy stocks listed than 2018. This indicates a potential volatility across this type of stock in general. The DQ stock in particular had an approximately 200% return in 2017 and a 63% loss in 2018 indicating significant risk to investing. 

Total volume is also reported in this table, which is a helpful indicator to performance. Generally, there is a correlation between the total volume and the stock price. When the trading volume is increasing, it will often indicate the stock price will trend upwards because of elevated interest. If the trading volume decreases, it may mean it is time to sell before there is a larger price reversal.(1)


** VBA Code Details**

The refactored code uses arrays to store the results for the trading volume, starting price, and ending price for each of the stocks. 
<p align="center" width="100%">
    <img width="33%" src=https://user-images.githubusercontent.com/105991478/175824523-478eda23-5710-4355-8c00-1a0bd69dfda1.png>
</p>


## References
(1) Nickolas, Steven. "Using Trading Volume to Understand Investment Activity." _Investors analyze trading volume when deciding whether to buy or sell a security,_ 01 Apr. 2022, https://www.investopedia.com/ask/answers/041015/why-trading-volume-important-investors.asp Accessed 26 June 2022.
