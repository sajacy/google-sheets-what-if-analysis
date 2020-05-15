What-If Analysis, Google Sheets Add-On
===========

Replicates the Excel "What-If Analysis" feature, **Data Tables**.

In the future, I may add the other Excel features, e.g. **Scenarios**.

### Install Add-On

This project is available as an Add-On in the Chrome Web Store:

[Install via Chrome Web Store](https://chrome.google.com/webstore/detail/what-if-analysis/plbdmkgogobniebdohfickjndnechhel)

### Manual Installation

Alternatively, you can copy the contents of the [`WhatIfAnalysis.gs`](WhatIfAnalysis.gs) file into a new script.

Step-by-step:

1. Within your Google Sheet, open the Script Editor (Tools > Script Editor)
2. Create a new script file, named WhatIfAnalysis.gs (File > New > Script file)
3. Paste the code from [`WhatIfAnalysis.gs`](WhatIfAnalysis.gs)
4. Save the script file (File > Save).
5. Reload your Google Sheet. 

### Example Sheet

After installation, make a copy of the [2018 Tax Calculator](https://docs.google.com/spreadsheets/d/1kQ8y3DsnCilScNda1AIcV_-M7TIAWSkdlF__ceZ-W7U/edit?usp=sharing) sheet as a working example.

### Usage

After installing the add-on, you should have a menu item **What-If Analysis > Create Data Table**.

This will try to create a data table from the currently-selected cell range.

For example, to create a 2-dimensional (bivariate) Data Table, you need to select a range
where the model output (formula) is the top-left cell, and where the two sets of values
are provided along the first row and first column.  The add-on will fill in the rest of the cells.

Before (example):

| **2D (bivariate) Data Table** |    |    |    |    |
| --------------- | -- | -- | -- | -- |
| rowinput       | 2  |    |    |    |
| colinput       | 10 |    |    |    |
| model output  (=rowinput+2*colinput):          | **22** | 10 | 20 | 30 |
|                 | 1  |  |  |  |
|                 | 2  |  |  |  |
|                 | 3  |  |  |  |
|                 | 4  |  |  |  |

After (example):

| **2D (bivariate) Data Table** |    |    |    |    |
| --------------- | -- | -- | -- | -- |
| rowinput       | 2  |    |    |    |
| colinput       | 10 |    |    |    |
| model output  (=rowinput+2*colinput):          | **22** | 10 | 20 | 30 |
|                 | 1  | 21 | 41 | 61 |
|                 | 2  | 22 | 42 | 62 |
|                 | 3  | 23 | 43 | 63 |
|                 | 4  | 24 | 44 | 64 |



To create a univariate Data Table, select 2 columns, with the model output (i.e. a formula; or a reference to a cell with the model output) in the first row, second-column (), and the test values in the first column.

Before (example):

| Univariate Data Table   |                     |
| ----------------------- | ------------------- |
|                         | input:              |
|                         | 50000               |
|                         | output (=input\*2): |
| Test input value column | **100000**              |
| 20000                   |                     |
| 50000                   |                     |
| 75000                   |                     |
| 100000                  |                     |
| 150000                  |                     |
| 200000                  |                     |
| 300000                  |                     |
| 500000                  |


After (example):
| Univariate Data Table   |                     |
| ----------------------- | ------------------- |
|                         | input:              |
|                         | 50000               |
|                         | output (=input\*2): |
| Test input value column | **100000**              |
| 20000                   | 40000               |
| 50000                   | 100000              |
| 75000                   | 150000              |
| 100000                  | 200000              |
| 150000                  | 300000              |
| 200000                  | 400000              |
| 300000                  | 600000              |
| 500000                  | 1000000             |

### Data Refresh

For performance reasons, Data Tables are not recalculated automatically.

Use the **What-If Analysis > Refresh Data Tables** menu item to refresh all data tables.

You can also delete data tables by deleting the associated Named Range (**Data > Named Ranges**), 
which will be named in the form of *DataTable_RxCy*.  The deleted cell range will no longer be 
considered a "Data Table" and you may freely edit / clear the cells as normal.

### Help - How to use WhatIfAnalysis.

Selected range must be at least 2 columns.
                               
**For a univariat Data Table (1 input variable):**
Top left cell can contain any text column header for the test values which need to follow downwards on the same leftmost column. The header for the second column needs to be a formula giving the model output (based on the input variable), or reference a cell with the model output.
      
**For a 2D (bivariat) Data Table (2 input variables):**
Top left cell needs to be a formula giving the model output (based on the input variables), or reference a cell with the model output. Furthermore, the leftmost column needs to have the various test values (for the row input variable) listed downwards. Also, the topmost row needs to have more than 1 column in addition to the leftmost column, and they need to contain test values (for the column input variable).
    
WhatIfAnalysis will in both cases fill in the remaining part of the matrix. It will insert test values into the cells which you specify as input variable(s), and fill in the matrix with the corresponding model output results.
    
### When to use WhatIfAnalysis, and when to use other approaches.
    
**For simple models then WhatIfAnalysis is not needed**, and results can be obtained much faster with repeating a formula normally throughout the matrix (in-spreadsheet formulas are generally faster than scripts like this WhatIfAnalysis script). For a bivariat Data Table (starting at top left cell A1) based on for instance a model like =(ROW_VARIABLE_CELL + COL_VARIABLE_CELL * 2), then the formula =$A2+B$1*2 could be input to B2 (cell 2,2 in the matrix) and dragged downwards to fill the entire first column, and then dragged to the right, to fill the entire matrix.
    
**But in more complex models** the output can be the result of a chain of formulas throughout various cells, and then WhatIfAnalysis is particularly useful. In some cases though, a conjoined formula can be reconstructed by going to the end output cell and into the formula there successively inserting the formulas contained in the cells it referenced. Then the approach mentioned in the simple model scenario above will work. This can give a significant speed increase (and allow for instantaneous and automatically refreshing values when changes occur), compared to using WhatIfAnalysis (which must be manually run to recalculate values). Although this approach is always theoretically possible, it is not always practically feasible (i.e. worth the effort). Typically in cases where several formulas in the chain of formulas specify ranges of cells which again include other formulas. Then WhatIfAnalysis comes handy.
    
**For the ultimately fastest approach with a complex model**, consider representing it entirely in code, and have it perform the calculations there (without relying on the spreadsheet). To do that, go to Tools -> Script editor, and write your own script using JavaScript. Your custom function in your script should take in all your input data (no formulas) and do all the calculations in code (instead of using spreadsheet formulas). This way the script can perform all the calculations in memory, without interacting with the spreadsheet at every step like WhatIfAnalysis does (which takes some time every time input values are set in the spreadsheet and output values are recorded from it). When your script has calculated all the outputs it can merely output them all to the spreadsheet at once (thus reducing the interactions with the spreadsheet to the minimum).
  


### License

MIT License

Copyright (c) 2018 Jeffrey Yang

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
