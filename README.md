# 2017 & 2018 Stock Analysis
#### The purpose of this analysis to provide Steve's parents with an idea of which stocks would be the best to invest it. 

### Analysis and Background:
  Originally Steve's parents were interested in DQ stocks. After using VBA to quickly pull in information for DQ we found this may not be the right fit for them. Since Steve's parents are still interested in investing in stock we decided to do some further coding so we could easily pull information for each of the stock options for both 2017 and 2018. Since we don't want our code to take too long incase more information on other stocks is ever added we had to go in and revamp some of the code in order to be sure it gives a quick and precise result. 
  
### Refactoring the code: 
  When we first wrote out our code it would run through the whole sheet to check the information each item we pulled (either the total daily volume or the return percentage). It was doing this in order to be sure we pulled the correct numbers needed for each ticker. This seems like it's not too big of a deal since we dont have _too_ much data but if we were to increase it by adding more stock options or more years this could take a lengthy time to run. You can see in our [original code](https://github.com/TorreyRawlings/stock-analysis/blob/main/UnfactoredCodeVisual.vb) we have a nested loop and nothing defining or reducing the criteria of what each line of code needs to check. In our [refactored code](https://github.com/TorreyRawlings/stock-analysis/blob/main/UnfactoredCodeVisual.vb) we added a _tickerindex_ in order to create a reference point for the code to go back to instead of reading through the entire sheet. Here is a visual of how much faster our code ran after doing the above:
### Original Code Run Time:
  ![Unfactored Code 2017](https://user-images.githubusercontent.com/103911529/167646828-a036f3c3-edca-4f98-9d3f-fa4aaf3f881f.png)
  ![unfactored code 2018](https://user-images.githubusercontent.com/103911529/167647059-4510e4a7-73e5-4ad0-acf6-1783812c954b.png)
### Refactored Code Run Time:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/103911529/167647146-c87e7570-c549-4de2-a847-3cc4ccd221b9.png)
![VBA-Challenge_2018](https://user-images.githubusercontent.com/103911529/167647182-f89299a4-6ae8-4e9e-911a-21795a8f88e1.png)

### Summary:
#### What are the benefits and disadvantages of refactoring code?:
  As you can see above refactoring the code was successful in being faster than the original code. Though, just because we _can_ refactor code though doesn't mean we always should or need to. If the dataset is small and has no plan to be referenced again or grow in size it might be easier if you are on a timeline to not refactor the code. It's usually a time consuming process that can also create room for error if you end up accidentally changing code that pulls the correct numbers. On the other hand for any code that you plan to reference/use multiple times or will be monitored by outside parties it's in your best interest to get the code to run as fast as you can so they aren't waiting on data any longer than they have to and there's room for the dataset to grow. This can also cause confusion for other developers who come in to view your code as it may not be as obvious what the code is trying to do.
#### What are the benefits and disadvantages of refactoring this specific code?:
  In this specific case scenario it took me a good few extra hours in order to refactor the code. There was also an issue where my results weren't coming out correctly to the original code I had which was frusterating and took even more time. It really only saved a few tenths of a second which in real time wouldn't feel like much to an end user. In this specific case I don't think the refactoring was generally worth the time it took me to complete BUT it would be worth it if we continue to add stock data into this file and it gets even more lengthy in size. 

