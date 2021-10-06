# Stock Analysis using VBA :chart:

## Overview of Project :mag:

### Purpose :dollar:
In this module we focused on visualizing and analyzing a series of market actions for Steve's parents, who wanted to invest in the stock market of companies that produce green energies. Using Microsoft's programming language called "Excel VBA" or "Visual Basic for Applications" a series of codes were made to guide the investor to make the best decision statistically in order to not only be able to make money by investing in the stock market, but also also help the environment by investing in these  green energy companies.

## Results :recycle:
In the stock market, 12 companies were studied throughout the years 2017 and 2018 to help the investors decide on which company should they invest. The companies studied were the following:
- Atlantica Yield ("AY")
- Canadian Solar ("CSIQ")
- Daqo New Energy ("DQ")
- Enphase Energy ("ENPH")
- First Solar ("FSLR")
- Hannon Armstrong ("HASI")
- Jinko Solar ("JKS")
- SunRun ("RUN")
- SolarEdge ("SEDG")
- SunPower ("SPWR")
- TerraForm Power ("TERP")
- Vivint Solar ("VSLR")

These were the faces of the green energies in the stock market on the years 2017-2018. The results of this analysis are mainly focused on the company to be studied and the comparison of the price of the tickers at the beginning of the year and the price at the end of the same year, the outcomes are the following:

### 2017 Stock Market :arrow_upper_right:
On this year the Stock Market showed the following results:
![](https://github.com/Frankdiazw/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

* **Figure 1. Stock Analysis for the 2017 Stock Market.**

In figure 1, we can observe the 12 companies to study, their "Total Daily Volumes" that represent the number of shares traded per day and the return in function of time, that represents the percentage of prices at the end of the year between the prices at the beginning of the year.

In this year, the results of the actions of the companies were satisfactory in most cases. We can see that for the company TerraForm ("TERP") it had a -7.2% return at the end of 2017, which means that in 2017 all those who invested in this company lost money and it was the only company that came out negative in that year. On the other hand, we have an impressive 199.4% return for Daqo New Energy ("DQ"), being the company that stood out the most this year, followed by SolarEdge ("SEDG") and Enphase Energy ("ENPH") with 184.5% and 129.5% respectively.

### 2018 Stock Market :arrow_lower_right:
On this year the Stock Market showed the following results:
![](https://github.com/Frankdiazw/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

* **Figure 2. Stock Analysis for the 2018 Stock Market.**

Now, the results this year were totally opposite to last year. It can be seen in the Return column that the majority came out negatively in 2018. We can see that Daqo New Energy ("DQ") was the company that fell the most with -62.6% compared to last year, which was the most outstanding company. However, only two companies came back positive by the end of 2018, Enphase Energy ("ENPH") and SunRun ("RUN") with 81.9% and 84.0% respectively.

For investors, they are recommended to invest their money in Enphase Energy ("ENPH") and SunRun ("RUN"), since they were the only companies that managed to raise their shares throughout the two years studied. It should not be forgotten that Daqo had an outstanding result in 2017, although it has decreased for 2018, another year would have to be reviewed to see if this company is as promising as in 2017. Comparing the two companies that managed to come out positive in the two years, Enphase Energy was the one that had the greatest impact with 129.5% and 81.9% on Return, recommending investors to invest in this company due to its statistical results.

### Description of code used to run the Stock Analysis :computer:
Click in the following link to see a detailed description of the code used in the workbook:
- :page_with_curl: [VBA code for Stock Analysis](https://github.com/Frankdiazw/Stock-Analysis/blob/main/VBA_Challenge.vbs)

## Summary :white_check_mark:
Talking about the refactored VBA code implemented on the Excel workbook, there are some advantages that can demonstrate that using a Refactored code is better. We can see in figure 1 and 2 the time lapses that show 4.89 second and 4.96 seconds on the refactored code, on the other hand we can observe in the following figures:
![](https://github.com/Frankdiazw/Stock-Analysis/blob/main/Resources/VBA_Original_Code_2017.png)

- **Figure 3. Time Lapse for the original code on the year 2017.**

![](https://github.com/Frankdiazw/Stock-Analysis/blob/main/Resources/VBA_Original_Code_2018.png)

- **Figure 4. Time Lapse for the original code on the year 2018.**

We can see in the Figures 3 and 4 that the codes ran slower due to the structure of the code, we can determine due to the difference of time lapses that the refactored code worked even better, and as if that weren't enough, the user is able to input the year he desires to analyze in the refactored code, in comparison with the original code were the user can't.

The disadvantages are presented when coding the refactored code, the programmer can possibly encounter with syntax errors and semantic errors. Personally, I encountered with the following errors:
- :heavy_exclamation_mark: [Application-defined or Object-defined error](https://stackoverflow.com/questions/17980854/vba-runtime-error-1004-application-defined-or-object-defined-error-when-select)
- :heavy_exclamation_mark: [Subscript out of range](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/subscript-out-of-range-error-9)
