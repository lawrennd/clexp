# clexp

Module for working with CL Expense Forms

This is a short script for aiding in creating expense statement forms for the computer lab.

Before you start, create a file in your working directory called `_config.yml`

It should be of the following form

```
title: # e.g. Professor, Dr etc
name: # Your full name
grant_code: # Grant code starting with GXXXX
four_digits: # Last four digits of your bank account number
payroll: # Your eight digit payroll reference number
```

Now you can install the package and create your expense report.

```python
%pip install git+https://github.com/lawrennd/clexp.git
```

You can download the template expense claim form from here.

```python
import urllib.request
urllib.request.urlretrieve('https://github.com/lawrennd/clexp/raw/main/template-expense-claims-partii.xlsx',
                           'template-expense-claims-partii.xlsx')
```

Now log in to your expensify account.

1.  Download the Expensify Report \'Expense Summary\' (top right corner
    on this link: <https://www.expensify.com/reports?param=>{}). Save it
    as `YYYY-MM-DD-bulk-export-id-expense-summary.csv` and update the
    `expense_data_csv` variable below  (use the real, year, date and month to keep track of things).

```python
expense_data_csv = 'YYYY-MM-DD-bulk-export-id-expense-summary.csv'
```

## Run the Script

Run the script below. It will create a separate excel spreadsheet
    for each report.

```python
import os
import sys
from datetime import date
import re
import pandas as pd

import clexp.expenses as exp

# Read the template file
form_df = pd.read_excel(os.path.join(exp.template_file), header=None)
exp.excel_style(form_df)
# Read the report data from Expensify
data_df = pd.read_csv(os.path.join(expense_data_csv), thousands=',', 
                      dtype={'original_amount': float, 'amount': float})

# Establish unique report ids.
report_ids = pd.unique(data_df.report_id)

# Write the reports
for i in report_ids:
    df2 = exp.write_claim(data_df[data_df.report_id==i], form_df)
```

## Submit the result

1.  Open `template-expense-claims-partii.xlsx`, select all (Ctrl + A).
    Open the report for the expenses. Do Paste special to paste the
    format onto the new excel spreadsheet.

    a. Paste the column widths first.
    b. Paste the format second.

2.  Go to Expensify and close the relevant report.

3.  Go to email, and 'save as' the PDF file from the report to
    your expenses directory (as specified in `_config.yml`) with file name that
    matches the excel file.

4.  Email the CL Accounts Team with the attached PDF and Excel
    spreadsheet.

