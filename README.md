# VBA-challenge
## Objective
This program analyzes stock data by year. It calculates the yearly change and its percent change in price for each stock entry, as well as the total stock volume.
## How it works
The program consists in 3 main parts:
- First the data is sorted in an ascending order based on the ticker symbol and the date. This ensures that the program will work properly even if the data is not in order.
- Then, the program starts analyzing each row focusing on the following variables: ticker, open value, close value and volume. The analysis starts with the open value of the first ticker. As long as the ticker value remains the same, the volume is being added. When the ticker changes, the close value of the old ticker is recorded and all values for the old ticker are printed in worksheet. Variables are initialized so the analysis of the next ticker can start. This process goes on until we reach the last row of the worksheet.
- Finally, the results are formatted. Once the program finishes the analysis of a particular ticker (stock), it prints the results and then it applies the following format:
  - if the yearly change in price is negative then the cell color changes to red. If it is equal or greater than 0 then it changes to green.
  - the percent change in price column is shown as %
 
