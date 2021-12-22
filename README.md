# VBA Green Stock Analyis

## Overview of Project

### Purpose

For Steve, I intially completed a series of macros in **VBA** which allowed him to quickly compile stock values by year.  The original ***All Stock Analysis*** utilizes nested loops to to roll through all rows in a chosen worksheet.  The loops sum or find values that corresponded to ticker values in the sheet.  To increase the usability of this macro and decrease the run time for the script, I've been asked to refactor the original code to eliminate the the nested loops by applying an index variable.  This report presents the results of the stock performance analysis; as well as, a description of results for the code refactoring.

## Results

### Analysis of Stock Performance for 2017 and 2018

There appears to be a drastic decrease in the majority of the stocks chosen for analysis between 2017 and 2018 (See Tables).  The best performing stock from **2017** and **2018** would be ***ENPH***.  ***ENPH*** had rate of returns of ***129.52% and 81.92%*** in 2017 and 2018 respectively and total volumes for each also rose substantialy between 2017 thru 2018 (**385,701,400** ).  The only other stock to have positive returns both years is ***RUN*** with ***5.55% and 83.95%*** in 2017 and 2018 respectively and the total volume for each also rose substantialy between 2017 thru 2018 (**235,075,800** second only ***ENPH***).  That said, despite negative returns in 2018; had someone purchased either ***DQ or SEDG*** in 2017 (**199.45% and 184.47% respectively**) they still would have siginficant returns at the end of 2018 even after losing ***62.6% and 7.75%*** respectively.  Total Volumes for ***DQ or SEDG*** still rose an additional 72,077,700 and 30,327,100 from 2017 to 2018.  On the negative side, ***TERP*** was the only stock to lose value both years; however, by trade volume ***SPWR and FSLR*** lost the most volume trading between 2017 and 2018 (-244,162,700 and -206,067,500).  The sharp decrease in trade volume for ***SPWR and FSLR*** could indicate more downside to these stocks recognized by the market.

![image](https://user-images.githubusercontent.com/91850824/146859267-ca0fda3d-8fea-4155-be6b-6878233e8396.png)  ![image](https://user-images.githubusercontent.com/91850824/146859281-330d0579-e7da-4e78-b525-ddb441371b4b.png)

### Code Refactoring

#### Original VBA yearValueAnalysis subroutine
The original script for the yearly stock analyisis can be found in module 1 in the **yearValueAnalysis** subroutine.  The Nested loops were utilized as the primary function for sorting and finding the ***Total Volumes and Starting/Ending Prices***.  It began with an array of ticker values. 

![image](https://user-images.githubusercontent.com/91850824/146860341-e0327bdb-c7bf-424c-bc17-941b6be5ecf2.png)

This array was used as the primary sorting item for the first "i" loop.   The first loop runs the script through the array index range **0 to 11**, and zeros out the ***totalVolume*** variable after each loop.  The next loop evaluates the ***ticker(i)*** over all rows in the first ticker column and sums the Values of the volume and finds the starting and ending prices corresponding to the array's index.  Conditional If Statements are used to chose values that match the corresponding ticker array.

![image](https://user-images.githubusercontent.com/91850824/146860794-f453dd73-1e7c-444c-bf6d-513cd555143f.png)

The loop, beginning with 0, progresses through the second loop and chooses values with ***tickers(0)***.  In this case, the ticker is ***"AY"***.  The script then closes the nested loop and  activates the **AllStocksAnalysis** report page to populate the cells in the report (***Ticker, Total Volume, and Return***).  The original loop range (0 To 11) is manipulated so that the calculated values can use the same range for the inputs.  

![image](https://user-images.githubusercontent.com/91850824/146861540-2631fa46-efd7-4f34-b410-bd7ba8f4c0d4.png)

The script then repeats over the entire data, but now using the next value in the loop range,index value 1 (***tickers(1)***), to calculate values for ***CSIQ***.  This is continued sequentially through index value 11

#### Refactored AllStockAnalysisRefactored subroutine

In contrast, the refactored script; which utilizes the same tickers array, solves the problem of data sorting by using a ticker index variable to assign the tickers array index in the script.  Taking this a step further, Arrays were created for the variables ***totalVolumes, starting prices, and ending prices*** as well.  

![image](https://user-images.githubusercontent.com/91850824/146862707-816004e5-d91c-45ee-8c19-b21c7ec62bd8.png)

Multiple loops are used (to initialize the tickerVolume index values to 0, to search all rows, and to inport final information); however, the nested loop was avoided.  

![image](https://user-images.githubusercontent.com/91850824/146862790-6cb3ae99-e2d6-4a11-9cbc-64f94e976043.png)

![image](https://user-images.githubusercontent.com/91850824/146862803-d6be84f4-a90f-486f-901f-15454ed88c12.png)

![image](https://user-images.githubusercontent.com/91850824/146862823-d884daaa-84d6-4a00-b662-19620dfcdcff.png)

By using a variable instead of a loop, the script assigns the index value for the ticker.  This eliminates time that would otherwise be spent loopping through the sheet again and again.  It also allows the values for each ticker (Volumes and Starting/Ending Prices) to be calculated and saved as values in each array and indexed to the same values used to index the tickers.  This enabled me retieve values for my variable arrays  based off the saved values recalled by their corresponding indicies.  The original script had to Activate a new sheet within each loop to report values, the refactored script can report the values in a seperate, more efficient loop to save time.  

The **all row loop**, now looks at each row of the sheet and gathers the data as it evaluates each row 1 at a time.  This saves a lot of time.  The original script had to evaluate every row for each of the 12 loops.  Now the refactored script evaluates each row and a conditional if statement assigns whether the tickers index should change.  With this logic, the variables can be updated in the appropriate index after reviewing each row once.  

![image](https://user-images.githubusercontent.com/91850824/146864369-4013bb38-86cb-4717-ac56-3faf5e59aac6.png)

#### Performance Original Versus Refactored

For both 2018 and 2017 runs, the refactored script was significantly faster.  A timer was included in the script to measure the starting and ending times during each scripts run.  The original **2017** run was completed in ***0.746 seconds***, the refactored run was completed in ***0.133 seconds***. 

![Module_2017_timer](https://user-images.githubusercontent.com/91850824/146864741-6197471f-4911-4a7e-87ea-df782b298c07.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91850824/146864753-50894705-9657-49ed-bc38-e4fd99e2c448.png)

The original **2018** run was completed in ***0.75 seconds***, the refactored run was completed in ***0.137 seconds***.

![Module_2018](https://user-images.githubusercontent.com/91850824/146864773-c7080019-3f0d-4957-ba73-75cddf257016.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/91850824/146864782-37955ab5-d889-461b-87e4-e29851d1f05b.png)

## Summary

### 1.) Advantages and Disadvantages of refactoring code?

Refactoring can greatly increase the time and memory required to run subroutines over large datasets.  The codes themselves can be made more efficient and smaller.  But, like any revision, it is easy to get lost in the changes and without solid orginization and documentation.  It may also be difficult to follow someones refactoring efforts if not aware of the original script.  Likewise, simplifying the subroutine for a specific objective may make that script only applicable to that objection and cannot be reused, or if editing is needed, more spots may require edits (i.e. index numbering for new array variables).  On a positive note, the refactored script saves more information to indicies so that more complicated calculations could be performed without adding more time to the loops. 

### 2.) How do these pros and cons apply to refactoring and the original VBA script?

The original script, was defintely slower; however, if we were to add additonal items to the tickers array we would only need to add the item and change the index total in the tickers array and the range in the first loop.  The refactored script would also require changes to all the other variable arrays as well as the ticker array.  

The refactored arrays save each value in a seperate index which gives you more functionality for computing other calculations later.  The original script requires the data to be looped through for the total number of indicies to create a new value to save, the indexed values in the refactor exist without being re-looped and imported.  

The drawback with both scripts, in comparison to our data set, is sorting.  Both rely on the data to be sorted by ticker and by price date.  If the sheet were not well sorted, the results, especially for the starting price and ending price could be drastically different.  The original script has an advantage over the the refactored script, in that, it searches the entire dataset for the ticker index each time.  So, even if the tickers are rearranged, the total volumes would still be accurately calculated even if the prices are solely relient on the date and ticker sorting.  

The refactored script is solely relient on the data's sorting.  Because the refactored script adds 1 to the ticker index if the ticker name changes, it will add ticker indicies beyond the defined indicies in the arrays and cause an error if the ticker names are all scrambled.  Both of these scripts would be benefitted with a routine that sorts the columns by ticker name and price date before the loops.  
