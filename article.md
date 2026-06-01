---
author: "Kyle Jones"
date_published: "August 23, 2025"
date_exported_from_medium: "November 10, 2025"
canonical_link: "https://medium.com/@kyle-t-jones/working-with-excel-data-in-python-960518f789a2"
---

# Working with Excel Data in Python

Excel is one of the most common tools for analysts. It's powerful, but cleaning data in Excel often involves endless formulas, manual...

### Working with Excel Data in Python
Excel is one of the most common tools for analysts. It's powerful, but cleaning data in Excel often involves endless formulas, manual edits, and a lot of clicking around. That process can be error-prone and hard to repeat.

Python offers a different path. With the Openpyxl library, you can treat Excel workbooks as programmatic objects --- read them, transform them with Pandas, and save the results back into a new sheet or file. This approach is faster, reproducible, and easier to maintain.

In this tutorial, we'll take a dataset from the *How to Clean Excel Data* exercise by Breaking Into Wall Street and show how to automate the cleaning process in Python. The same transformations you would do by hand in Excel formulas can be scripted in a few lines of code.

### Why Use Python Instead of Excel?
When you replace manual Excel cleaning with Python, you gain several advantages:

1.  [Speed --- Pandas processes data faster than Excel formulas.]
2.  [Replicability --- You can apply the same cleaning steps across new datasets without starting over.]
3.  [Readability --- Python code is easier to read than long, nested formulas.]
4.  [Automation --- You can run the script automatically and avoid human error.]

### Step 1: Install and Import Libraries
```python
! pip install openpyxl
import pandas as pd
import openpyxl
```

### Step 2: Load the Workbook and Data
``` 
filename = 'data/XL-02-PC-05-Cleaning-Up-Data-Before.xlsx'

wb = openpyxl.load_workbook(filename)
ws = wb['Data']
df = pd.DataFrame(ws.values)
```

### Step 3: Fix the Headers
```python
def make_header(df: pd.DataFrame)-> pd.DataFrame:
    df.columns = df.iloc[0]
    df.drop(df.index[0], inplace=True)
    return df

df = make_header(df)
```

### Step 4: Split Columns
Instead of using Excel formulas to split text, Pandas can do it directly.

``` 
df[['First Name', 'Last Name']] = df['Customer Name'].str.split(" ", n=1, expand=True)
df[['Address', 'City', 'State', 'ZIP']] = df['Address, City, State, and ZIP'].str.split(",", n=3, expand=True)
```

### Step 5: Standardize Text
``` 
df['Address'] = df['Address'].str.title()
df['City'] = df['City'].str.title()
df['State'] = df['State'].str.upper()
```

### Step 6: Save Cleaned Data Back to Excel
```python
from openpyxl.utils.dataframe import dataframe_to_rows

wb.create_sheet(index=0, title='Cleaned Data') 
ws = wb['Cleaned Data']
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)
filename_after = 'data/XL-02-PC-05-Cleaning-Up-Data-After.xlsx'
wb.save(filename_after)
wb.close()
```

### Before and After
Here's what the data looks like before and after cleaning:


The data cleaning tool in Excel has improved a lot, but I still find it easier to use Python if there is a lot of cleanup to do.

### Resources
- [[Openpyxl documentation](https://openpyxl.readthedocs.org/en/default)]
- [[Automate the Boring Stuff with Python](https://automatetheboringstuff.com/)]

### Summary
In this example, we used Python and Openpyxl to replicate what's usually a manual cleaning exercise in Excel. By splitting names, parsing addresses, and standardizing text, we automated a process that normally takes multiple formulas and manual edits.

Python doesn't replace Excel --- it makes Excel better. It turns spreadsheets into reliable inputs and outputs in a repeatable workflow, giving analysts time back to focus on insights rather than formatting.
