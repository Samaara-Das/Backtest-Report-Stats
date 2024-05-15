# What this application does
This extracts data from HTML backtest reports and adds that data into a single Excel file. This doesn't have a UI.

# Requirements
2. Execute this application from main.py.

## To-dos in main.py:
1. `SHARED_FOLDER_PATH` should be the path to a shared folder which contains HTML backtest reports.
2. `SCHEDULER_TIME_FRAME` should be the interval at which new backtest reports are processed. 

## To-dos in excel_utils.py
1. `EXCEL_FILE_PATH` should be the path to the Excel file which will contain data from all the backtest reports.