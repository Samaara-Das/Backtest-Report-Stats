import os
import glob
import logger_setup
import pandas as pd

# Set up logger for this file
main_logger = logger_setup.setup_logger(__name__, logger_setup.logging.DEBUG)

if __name__ == '__main__':
    try:
        path = os.path.expanduser('~/Desktop/Backtest Reports') # Set the path to the "Backtest Reports" folder on the Desktop
        os.chdir(path) # Change the directory to the specified path
        excel_files = glob.glob('*.xls*')  # This will capture both .xls and .xlsx files which are assumed to be backtest reports

        for file in excel_files: 
            df = pd.read_excel(file) # Load the Excel file into a pandas DataFrame
            
            # Save the DataFrame back to Excel without any formatting
            # You can specify 'engine='openpyxl'' if you encounter any issues
            df.to_excel(file, index=False)
            
            print(f"Processed {file} - formatting removed.")

    except Exception as e:
        main_logger.error(f"Error: {e}")
        raise e
