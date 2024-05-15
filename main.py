import os
import time
import logger_setup
import browser
import schedule
from excel_utils import setup_excel_file, EXCEL_FILE_PATH
from backtest_reports_processor import process_html_file, titles_and_xpaths

REMOVE_LOG = True

# Set up logger for this file
main_logger = logger_setup.setup_logger(__name__, logger_setup.INFO)

# specify the path to the shared folder which has all the HTML reports
SHARED_FOLDER_PATH = "D:\\Shared folder of HTML Reports"

# Clean up the log
if REMOVE_LOG:
    with open('app_log.log', 'w') as file:
        pass

def get_current_backtest_reports(folder_path):
    """
    Gets HTML backtest reports in the specified folder.

    Args:
    - folder_path (str): The path to the folder.

    Returns:
    - set: A set of file paths.
    """
    try:
        return {os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".htm") or f.endswith(".html")}
    except Exception as e:
        main_logger.error(f"Error getting current backtest reports from {folder_path}: {e}")
        raise

def check_for_new_backtest_reports():
    """
    Check for new backtest reports in the shared folder and process them.
    """
    try:
        current_files = get_current_backtest_reports(SHARED_FOLDER_PATH)
        new_files = current_files - processed_files

        # Process new files
        main_logger.info(f"New files: {new_files}")
        for file_path in new_files:
            process_html_file(file_path, browser)
            processed_files.add(file_path)
            main_logger.info(f"Processed new file: {file_path}")

    except Exception as e:
        main_logger.error(f"Error checking for new files: {e}")
        raise

# Run main code
if __name__ == '__main__':
    try:
        # initiate Browser
        browser = browser.Browser(keep_open=True, headless=True)

        # Setup the Excel file
        setup_excel_file(EXCEL_FILE_PATH, titles_and_xpaths)

        # Process all the existing backtest reports
        processed_files = get_current_backtest_reports(SHARED_FOLDER_PATH)
        for file_path in processed_files:
            process_html_file(file_path, browser)

        # Schedule the check for new backtest reports every hour
        main_logger.info("Scheduling check for new backtest reports...")
        schedule.every(1).hour.do(check_for_new_backtest_reports)

        while True:
            schedule.run_pending()
            time.sleep(1)  # Wait a bit before checking the schedule again

    except Exception as e:
        main_logger.exception(f'Error in main.py: {e}')

    finally:
        browser.driver.close()
