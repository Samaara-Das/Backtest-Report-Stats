## Branch Info
This branch is for extracting data from back test reports into 1 excel file on the local computer. This doesn't have a UI. Execute this application from main.py.

## To-dos in main.py:
1. `SHARED_FOLDER_PATH` should be the path to shared folder which contains the HTML backtest reports.
2. `SCHEDULER_TIME_FRAME` should be the interval at which new backtest reports are processed. 

## To-dos in excel_utils.py
1. `EXCEL_FILE_PATH` should be the path to the Excel file which will contain data from all the backtest reports.