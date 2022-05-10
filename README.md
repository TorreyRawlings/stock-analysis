# 2017 & 2018 Stock Analysis
#### The purpose of this analysis to provide Steve's parents with an idea of which stocks would be the best to invest it. 

### Analysis and Background:
  Originally Steve's parents were interested in DQ stocks. After using VBA to quickly pull in information for DQ we found this may not be the right fit for them. Since Steve's parents are still interested in investing in stock we decided to do some further coding so we could easily pull information for each of the stock options for both 2017 and 2018. Since we don't want our code to take too long incase more information on other stocks is ever added we had to go in and revamp some of the code in order to be sure it gives a quick and precise result. 
  
### Refactoring the code: 
  When we first wrote out our code it would run through the whole sheet to check the information each item we pulled (either the total daily volume or the return percentage). It was doing this in order to be sure we pulled the correct numbers needed for each ticker. This seems like it's not too big of a deal since we dont have _too_ much data but if we were to increase it by adding more stock options or more years this could take a lengthy time to run. You can see in our original code we have a nested loop and nothing defining or reducing the criteria of what each line of code needs to check. In our refactored code we added a _tickerindex_ in order to
