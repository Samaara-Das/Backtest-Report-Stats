import logger_setup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, TimeoutException

# Set up logger for this file
main_logger = logger_setup.setup_logger(__name__, logger_setup.logging.DEBUG)

CHROME_PROFILE_PATH = 'C:\\Users\\Puja\\AppData\\Local\\Google\\Chrome\\User Data'

class Browser:

    def __init__(self, keep_open: bool, headless: bool) -> None:
        chrome_options = Options() 
        if headless:
            chrome_options.add_argument('--headless') # the application will run without opening the chrome browser and be lightwight on system resources. This also won't interfere with other selenium controlled browsers.
        chrome_options.add_experimental_option("detach", keep_open)
        chrome_options.add_argument('--profile-directory=Profile 2')
        chrome_options.add_argument(f"--user-data-dir={CHROME_PROFILE_PATH}")
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        main_logger.info('Browser initialized')

    def open_page(self, url: str):
        '''This opens `url` and maximizes the window'''
        try:
            self.driver.get(url)
            self.driver.maximize_window()
            main_logger.info(f'Opened {url}')
            return True
        except WebDriverException:
            main_logger.exception(f'Cannot open this url: {url}. Error: ')
            return False 