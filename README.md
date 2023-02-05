# Trial-Balance-comparisons
This is a VBA code that helps in merging two trial balances and creates a comparison. Proven helpful for the accountants
The code performs the following tasks:
Declares variables for the worksheets, the last row in each worksheet, and a flag for indicating a match.
Initializes the worksheet variables with the corresponding worksheets in the current workbook.
Gets the last row in each of Sheet1 and Sheet2 by finding the last populated cell in column A.
Loops through all rows in Sheet1, starting from row 2. For each row, it checks if the value in column A of the current row in Sheet1 matches the value in column A of any row in Sheet2. If a match is found, it sets the flag to true and stores the values from both sheets in the corresponding columns in Sheet3. If no match is found, it sets the interior color of column A in Sheet3 to red to indicate an error.
Loops through all rows in Sheet2, starting from row 2. For each row, it checks if the value in column A of the current row in Sheet2 matches the value in column A of any row in Sheet1. If no match is found, it stores the values from Sheet2 in the corresponding columns in Sheet3, with a 0 in column C, and sets the interior color of column A in Sheet3 to red to indicate an error.
Note: The code uses the xlUp constant from the Excel object library to find the last row in a column.
