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

def get_current_backtest_reports(source_folder):
    """
    Gets HTML backtest reports in the specified folder.

    Args:
    - source_folder (str): The path to the folder.

    Returns:
    - set: A set of file paths.
    """
    try:
        return {os.path.join(source_folder, f) for f in os.listdir(source_folder) if f.endswith(".htm") or f.endswith(".html")}
    except Exception as e:
        main_logger.error(f"Error getting current backtest reports from {source_folder}: {e}")
        raise

def check_for_new_backtest_reports(source_folder, processed_files):
    """
    Check for new backtest reports in the shared folder and process them.
    """
    try:
        current_files = get_current_backtest_reports(source_folder)
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

def run_main_script(source_folder, destination_file, stop_event, time_frame):
    """
    Main script to process backtest reports.

    Args:
    - source_folder (str): The path to the source folder containing HTML reports.
    - destination_file (str): The path to the destination Excel file.
    - stop_event (threading.Event): Event to signal stopping the script.
    - time_frame (int): Time frame in minutes for scheduling the checks.
    """
    try:
        # Initiate Browser
        browser_instance = browser.Browser(keep_open=True, headless=True)

        # Setup the Excel file
        setup_excel_file(destination_file, titles_and_xpaths)

        # Process all the existing backtest reports
        processed_files = get_current_backtest_reports(source_folder)
        for file_path in processed_files:
            if stop_event.is_set():
                main_logger.info("Script execution stopped.")
                return
            process_html_file(file_path, browser_instance)

        # Schedule the check for new backtest reports
        schedule.every(time_frame).minutes.do(check_for_new_backtest_reports, source_folder, processed_files)

        while not stop_event.is_set():
            schedule.run_pending()
            time.sleep(1)  # Wait a bit before checking the schedule again

    except Exception as e:
        main_logger.exception(f'Error in main.py: {e}')

    finally:
        browser_instance.driver.close()
