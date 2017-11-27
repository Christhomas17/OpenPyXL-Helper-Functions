# OpenPyXLHelper-Functions
Some simple functions that adapt OpenPyXl to fit my needs


OpenPyXl lets you easily manipulate Excel files using Python but it is not made to work with R1 ranges such Game!C13:E27 but these are the types of ranges that I am accustomed to so I made my own function, which to the best of my knowledge, seems to handle most most issues.
So for example: 
```
from openpyxl import load_workbook
wb = load_workbook('Math.xlsx', data_only = True)
df = data_from_range(r"Game!C13:E27",wb)

```

will return a dataframe with the values contained in that range. The data_only allows you to get the values rather than the formulas in the cells. 

I also created a simple function that will allow to write a dataframe to a specific range in an Excel sheet. Pandas has a buuilt in function to write a dataframe to a workbook but I could not figure out how to get it to write to a specific range, thus this function. 

Using the df created from above, you can simply:
```
data_to_range(df,r"Game!C13:E27",wb)
```
