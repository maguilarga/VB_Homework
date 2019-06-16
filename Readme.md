# Stock Data Summary

Prepare a VB code that will create a summary of stocks traded once is run.

## Getting Started

Copy the XLSM file into a local directory

### Prerequisites

Excel

## Running the tests

1. Open the file with Excel. 
2. Execute one of the macros (StockSummary, StockSummary_Array)

The program will execute in each worksheet and create a summary. The summary will be located uin the same worksheet starting on column J


### And coding style tests

The requirement was solved in three different ways:
* Summ_Stock.vbs - in this case, coding techniques presented during class were used to solve the problem at hand.
* Summ_Stock_Array.vbs - during my research for the homework, I came across some information about reading data into memory arrays. I found the idea interesting so I put it to work. The code executes really quickly, but would need to be tested for extremly big data sets (1 million rows) in Excel; it might not be able to handle all data in memory.
* Summ_Stock_Range.vbs - during my research for the homework, I came across some information about how to define subrange out of a bigger range of data. The code works correctly but executes fairly slowly, that is the reason I did NOT bring this code to the final XLSM file.


## Built With

* Excel
* Visual Basic


## Versioning

v 1.0


## Authors

* **Martha E. Aguilar** - *Initial work*
