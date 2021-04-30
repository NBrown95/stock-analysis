# Refactoring Code

## Overview/Purpose of Project

For this project, we are refactoring code analyzing stock data. We want this VBA script to be able to efficiently analyze thousands of stock data. To do this, we must refactor the code to loop through all of the data one time.

## Results

### Execution Time
The original code used a nested for loop to cycle through the stock tickers and then each row to find the cell we are focused on. The following images show the time it takes to run this analysis for 2017 and 2018 respectively.

![Screenshot 1](https://user-images.githubusercontent.com/81498850/116643345-7c14d400-a936-11eb-9886-5aeaa014e560.png)

![Screenshot 2](https://user-images.githubusercontent.com/81498850/116643389-9353c180-a936-11eb-9343-56351bcc5a3a.png)

The refactored code created a new variable and new output arrays that allow analysis of tickers and their respective rows in one action. This allows for the run time to be significantly faster. The run times for 2017 and 2018 for the refactored code is shown in the following two images. As we can see, this new refactored code allows for far greater speeds. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81498850/116643587-ff362a00-a936-11eb-9d0b-c4d5a5a9dae5.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/81498850/116643634-19700800-a937-11eb-82d2-ab8a81607a53.png)

### Stock Performance

As for the stocks themselves, it is clear to see that in 2017, eleven of the twelve stocks saw a positive return with DQ, ENPH, FSLR, and SEDG more than doubled their return. This is seen in the following image.

![Screenshot 3](https://user-images.githubusercontent.com/81498850/116643453-b7170780-a936-11eb-8b11-bb8f0903f077.png)

However, in 2018, only two of the twelve stocks saw a positive return. Interestingly, ENPH was the only stock to have a positive return for both years. The performance for these stocks in 2018 can be seen below.

![Screenshot 4](https://user-images.githubusercontent.com/81498850/116643478-c9914100-a936-11eb-8e83-124b11b2346d.png)

## Summary

### Advantages/Disadvantages of Refactoring Code

The advantages of refactoring code are that we are able to perform the same functions as before in fewer steps. This allows the code to run faster and more efficiently than before. In addition to this, refactoring code will use less memory and make it easier to read for other people who may be using this code. The disadvantage of refactoring is that it takes time and effort to determine the proper logic and loops to refactor correctly.

### Advantages/Disadvantages of This VBA Script

Specifically, with this VBA script, the advantages include decreasing the amount of time it takes to loop through twelve stocks and analyze their volumes and returns. If someone else were to use this code to analyze similar stocks, it would benefit them to understand the logic and to simplify the code as much as possible. With this refactored script, the code is not repeating itself and runs efficiently. Therefore, it can be well understood by others as well as the user.
