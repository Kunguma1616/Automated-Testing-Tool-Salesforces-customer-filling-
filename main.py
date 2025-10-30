import os # to get the environment varviable 
import time # which help us to pause  the script for a set of number 
import re  # regular expression operations it is used for advanced string matching and manipulation like normalizing psotcodes 
import pandas as pd # to read the salesforces data csv 
import logging # this is logging meassage to consosle and to log files 
import traceback  # use to get the detail information about the eroor 
import json  # use to save the final report as an .jscon file 
import zipfile  # to create the zip file 
import random # <--- ADDED FOR RANDOM DIVISION SELECTION
from pathlib import Path # easier way to handle file and floder
from datetime import datetime  # uses to get the current data and time . the script uses the current dat and time for folder name  
from selenium import webdriver # this is the main component that starts and controls the web browser 
from selenium.webdriver.common.by import By # providing the methods for finding elements on the web page such as xpath 
from selenium.webdriver.support.ui import WebDriverWait, Select # this tool tells seleinuim to wait fo a certain condition instead of faling immidiately
from selenium.webdriver.support import expected_conditions as EC # provides the condition that webdriver waits for an elsemnt to exsit in the html 
from selenium.webdriver.chrome.service import Service # seliliuim talk to the chrome 
from selenium.webdriver.chrome.options import Options # keep it open after the script finsishes for the chrome browser 
from selenium.webdriver.common.keys import Keys # allow the scripte to simulatee pressing keyword keys 
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager # a very hepful ulititlity it automatcallly downlads and manages the correct version of the chromeve .exe to match the chrome broswer 
from dotenv import load_dotenv # read the .env file from yoour scrpite directory and load the variable inside into the environment 
from PIL import Image  # Added as requested , it is python image library 



# =========================
# ROBUST .ENV LOADING
# =========================
try:
    script_dir = Path(__file__).parent
    env_path = script_dir / '.env'
    if env_path.exists():
        load_dotenv(dotenv_path=env_path)
        print(f"✓ Credentials loaded successfully from {env_path.resolve()}")
    else:
        print(f"   WARNING: .env file NOT FOUND at {env_path.resolve()}")
        print(f"   Create a .env file with your credentials or set environment variables.")
except Exception as e:
    print(f" Error loading .env file: {e}")

# =========================
# CONFIGURATION
# =========================
# SharePoint configuration has been REMOVED.

# Salesforce Configuration
LOGIN_URL = "https://test.salesforce.com/"
HOME_URL = "https://chumley--staging.sandbox.lightning.force.com/lightning/page/home"
SF_USERNAME = os.getenv("SF_USERNAME", "kunguma.balaji@aspect.co.uk.staging")
SF_PASSWORD = os.getenv("SF_PASSWORD")
CSV_FILE = r"C:\Users\User\Downloads\salesforce_synthetic_data.csv"
DIVISIONS = ["Homeowner", "Insurance", "Corporate", "Government"]
DEFAULT_DIVISION = os.getenv("DEFAULT_DIVISION", DIVISIONS[0])
# Check for missing credentials
if not SF_PASSWORD:
    SF_PASSWORD = input("Balaji@1616")

# =========================
# SHAREPOINT UPLOADER CLASS (REMOVED)
# =========================


# =========================
# AUTOMATION REPORTER CLASS
# =========================
class AutomationReporter:
    """Logs steps, takes screenshots, and manages local artifacts."""
    
    def __init__(self, driver, base_run_name="Automation_Run"):
        self.driver = driver # it takes the seliliuim driver and saves it so it can be used later to take screenshort 
        self.step_counter = 0 # it create the counter varibale and stats it at 0 
        self.run_timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") # it excalty get the dat and time downd to the second and formats it into a string 
        self.run_name = f"{base_run_name}_{self.run_timestamp}"# it combine the base run ame ike automation run witht he timestamp from the line examp;e automation dtaa nd time 
        
        # Create local artifact directory
        self.local_artifact_dir = Path("./artifacts") / self.run_name # it figure it part out the full pth for our new run folder 
        self.local_artifact_dir.mkdir(parents=True, exist_ok=True) # it will create the floder o our computer
        
        self.log_file_path = self.local_artifact_dir / "automation_run.log" 
        self.report_json_path = self.local_artifact_dir / "report.json" # it will figure 
        
        self.report_data = [] # create a new empyty list memeory 
        
        # Set up file handler for this specific run
        self.logger = logging.getLogger(f'AutomationReporter_{self.run_timestamp}') 
        self.logger.setLevel(logging.INFO)
        self.logger.handlers.clear()  # Clear any existing handlers
        
        file_handler = logging.FileHandler(self.log_file_path)
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(file_handler)

        logging.info(f" Artifact directory: {self.local_artifact_dir.resolve()}")

    def _log(self, level, step_name, message, exception=None):
        """Internal logging method."""
        log_time = datetime.now().isoformat()
        
        # Clean step name for filenames
        safe_step_name = re.sub(r'[^\w\-_\. ]', '_', step_name).strip()
        safe_step_name = re.sub(r'\s+', '_', safe_step_name).lower()
        
        self.step_counter += 1
        filename = f"{self.step_counter:03d}_{safe_step_name}.png"
        filepath = self.local_artifact_dir / filename
        
        # 1. Take Screenshot
        try:
            self.driver.save_screenshot(str(filepath))
        except Exception as e:
            logging.error(f"Failed to take screenshot: {e}")
            filename = None

        # 2. Log to file
        log_message = f"STEP {self.step_counter:03d} [{step_name}]: {message}"
        if level == "ERROR":
            self.logger.error(log_message)
            if exception:
                self.logger.error(f"EXCEPTION: {exception}")
                self.logger.error(traceback.format_exc())
            logging.error(log_message)
        else:
            self.logger.info(log_message)
            logging.info(log_message)

        # 3. Append to JSON report data
        self.report_data.append({
            "step": self.step_counter,
            "timestamp": log_time,
            "status": level,
            "step_name": step_name,
            "message": message,
            "screenshot": filename,
            "exception": str(exception) if exception else None
        })

    def log_step(self, step_name, message):
        """Logs a successful step and takes a screenshot."""
        self._log("INFO", step_name, message)

    def log_error(self, step_name, message, exception=None):
        """Logs a failed step and takes a screenshot."""
        self._log("ERROR", step_name, message, exception)

    def save_report(self):
        """Saves the structured JSON report to disk."""
        try:
            with open(self.report_json_path, 'w') as f:
                json.dump(self.report_data, f, indent=4)
            logging.info(f" JSON report saved: {self.report_json_path}")
        except Exception as e:
            logging.error(f" Failed to save JSON report: {e}")

    def create_artifacts_zip(self):
        """Saves the final report and zips all artifacts."""
        # 1. Save final JSON report before zipping
        self.save_report()
        
        # Define zip file path (e.g., ./artifacts/Automation_Run_2025-10-28_11-30-00_artifacts.zip)
        zip_file_path = self.local_artifact_dir.parent / f"{self.run_name}_artifacts.zip"
        logging.info(f" Creating zip file: {zip_file_path}")
        
        try:
            with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Add all files from the artifact directory to the zip
                for file_path in self.local_artifact_dir.glob('*.*'):
                    zipf.write(file_path, arcname=file_path.name)
            
            logging.info(f" Successfully created artifacts zip: {zip_file_path.resolve()}")
            logging.info(f"   All screenshots, logs, and reports are inside this zip file.")
            
        except Exception as e:
            logging.error(f" Failed to create zip file: {e}")
            logging.error(traceback.format_exc())
            logging.info(f"   Artifacts are still available in the folder: {self.local_artifact_dir.resolve()}")

    # upload_artifacts_to_sharepoint method has been REMOVED.


# =========================
# DIVISION / SECTOR / BUSINESS CONFIGURATION
# =========================
ALLOWED_DIVISIONS = ["Homeowner", "Insurance", "Corporate", "Government"]

DIVISION_TREE = {
    "Homeowner": {
        "Homeowner / occupier": ["Homeowner / occupier"],
        "Landlord (multiple properties)": ["Landlord (multiple properties)"],
    },
    "Insurance": {
        "Utilities": ["Utilities"],
        "Commercial": ["Commercial"],
        "Domestic": ["Domestic"],
    },
    "Corporate": {
        "Education": ["College", "Preschool", "University", "School", "Others"],
        "Healthcare": ["Care home", "Vet", "Health centre", "Hospital", "Pharmacy", "Opticians", "GP"],
        "Food & beverage": ["Fast food", "Restaurants", "Bar", "Cafe", "Commercial food", "Private members pub"],
        "Charity": ["Office and shops", "Office", "Charity shop", "Others"],
        "Religious buildings": ["Others", "Church", "Mosque", "Religion centre", "Temple"],
        "Retail": ["Sports", "Travel agent", "Bank", "Book shop"],
        "Office": ["Others", "Business park", "Business club", "Private office"],
        "Leisure & entertainment": ["Bingo club", "Bowling", "Casino", "Museum", "Theatre"],
        "Sports & fitness": ["Fitness studio", "Golf club", "Gym", "Kids play", "Sport club"],
        "Hotels": ["Guest houses", "Hostel", "Hotel", "Private member club", "Bed and breakfast"],
        "Property": ["Estate agent", "Maintenance contractors", "Housing association"],
        "Agriculture": ["Farm", "Horticultural", "Pet daycare"],
        "Manufacturing": ["Other", "Factory"],
        "Key accounts": ["Gym", "Business park", "Cafe", "Fast food"],
    },
    "Government": {
        "Foreign government": ["Other", "Diplomatic residence", "Embassy / high commission"],
        "Services": ["Other", "Community centre / youth centre", "Leisure", "Library", "Park / recreation",
                     "Refuse collection / processing", "Swimming pool"],
        "Housing": ["Others", "Council house"],
        "Courts": ["Other", "Court"],
        "Detention": ["Other", "Adult prison / youth prison"],
        "Military": ["Other", "Air force / army / navy"],
        "Emergency services": ["Other", "Ambulance", "Coast guard", "Fire", "Police", "Training facility"],
        "NHS": ["Other", "Care home / hospice", "Health centre / clinic", "Hospital", "Pharmacy", "GP / dental practice"],
        "Government / council offices": ["Other", "Government / council office"],
    },
}

PREFERRED_IDS = {
    "Division": ["lwc-6aa2906q6hj", "lwc-3anllt3fi9m", "lwc-184g1ot50fv-host"],
    "Sector Type": ["sectorType", "lwc-sector-type", "lwc-3anllt3fi9m"],
    "Business Type": ["businessType", "lwc-business-type", "lwc-184g1ot50fv-host"],
}


# =========================
# LOGGING SETUP
# =========================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automation_main.log'),
        logging.StreamHandler()
    ]
)


# =========================
# HELPER FUNCTIONS
# =========================
def get_row_value(row: pd.Series, keys):
    """Return the first non-empty value from row using any alias in keys."""
    for k in keys:
        if k in row and pd.notna(row[k]):
            val = str(row[k]).strip()
            if val != "":
                return val
    return ""


def initialize_driver():
    """Initialize Chrome WebDriver with automatic driver management"""
    try:
        options = Options()
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option("detach", True)

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logging.info("ChromeDriver initialized successfully")
        return driver
    except Exception as e:
        logging.error(f"Failed to initialize ChromeDriver: {str(e)}")
        raise


def login_to_salesforce(driver, reporter: AutomationReporter):
    """Login to Salesforce with proper authentication wait"""
    try:
        logging.info(" Navigating to Salesforce login page...")
        driver.get(LOGIN_URL)
        reporter.log_step("Navigate_Login", f"Navigated to {LOGIN_URL}")

        wait = WebDriverWait(driver, 30)

        # Enter username
        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        reporter.log_step("Found_Username_Field", "Found username field, about to enter text")
        username_field.clear()
        username_field.send_keys(SF_USERNAME)
        reporter.log_step("Enter_Username", "Username entered")

        # Enter password
        password_field = driver.find_element(By.ID, "password")
        reporter.log_step("Found_Password_Field", "Found password field, about to enter text")
        password_field.clear()
        password_field.send_keys(SF_PASSWORD)
        reporter.log_step("Enter_Password", "Password entered")

        # Click login
        login_button = driver.find_element(By.ID, "Login")
        reporter.log_step("Found_Login_Button", "Found login button, about to click")
        login_button.click()
        reporter.log_step("Click_Login", "Login button clicked")

        authenticated = False

        # Check if verification code page appears
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Verify Your Identity') or contains(., 'verification code')]"))
            )
            logging.warning("  VERIFICATION CODE REQUIRED!")
            logging.warning("   Please enter the verification code manually in the browser")
            reporter.log_step("Manual_Verification_Required", "Waiting 40s for manual verification")
            time.sleep(40)
            authenticated = True
        except TimeoutException:
            pass

        # Wait for Lightning/Home page
        if not authenticated:
            try:
                logging.info(" Waiting for Salesforce Lightning page...")
                wait.until(lambda d: "lightning" in d.current_url or "setup" in d.current_url)
                authenticated = True
                logging.info(" Login successful!")
                reporter.log_step("Login_Success", f"Login successful. URL: {driver.current_url}")
            except TimeoutException:
                pass

        if authenticated:
            time.sleep(5)
            return True
        else:
            logging.error("Login failed")
            reporter.log_error("Login_Failed", "Could not reach home page")
            return False

    except Exception as e:
        logging.error(f" Login failed: {str(e)}")
        reporter.log_error("Login_Exception", "Login failed with exception", e)
        return False


def switch_into_relevant_iframe(driver):
    """Try to switch into the iframe that contains the radio buttons."""
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        logging.info(f"Found {len(iframes)} iframe(s)")

        for idx, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)
                radios = driver.find_elements(By.XPATH, "//input[@type='radio']")
                if radios:
                    logging.info(f" Switched to iframe {idx} with {len(radios)} radio button(s)")
                    return True
            except Exception:
                continue

        driver.switch_to.default_content()
        return False
    except Exception as e:
        logging.warning(f"Iframe switch failed: {e}")
        driver.switch_to.default_content()
        return False


def navigate_to_form(driver, reporter: AutomationReporter):
    """Navigate to the form and select Create Domestic Customer"""
    try:
        wait = WebDriverWait(driver, 30)

        logging.info(f"Navigating to home page: {HOME_URL}")
        driver.get(HOME_URL)
        time.sleep(5)
        reporter.log_step("Navigate_Home", f"Navigated to {HOME_URL}")

        switched = switch_into_relevant_iframe(driver)
        if switched:
            reporter.log_step("Iframe_Switch", "Switched into form iframe")
        else:
            reporter.log_step("Iframe_Switch", "Proceeding in default content (no iframe needed)")
            
        time.sleep(2)

        logging.info(" Attempting to click 'Create Domestic Customer' radio...")
        radio_clicked = False
        domestic_radio_element = None

        domestic_elements = driver.find_elements(By.XPATH, "//*[contains(., 'Create Domestic Customer')]")
        logging.info(f"   Found {len(domestic_elements)} elements with 'Create Domestic Customer' text")

        # Strategy 1: Find parent with radio
        for elem in domestic_elements:
            try:
                parent_with_radio = elem.find_element(By.XPATH, "./ancestor::*[.//input[@type='radio']][1]")
                radio = parent_with_radio.find_element(By.XPATH, ".//input[@type='radio']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio)
                time.sleep(0.3)
                reporter.log_step("Found_Domestic_Radio_S1", "Found radio (Strategy 1), about to click")
                try:
                    radio.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", radio)
                logging.info(" Create Domestic Customer radio clicked (Strategy 1)")
                reporter.log_step("Click_Domestic_Radio", "Clicked 'Create Domestic Customer' radio")
                radio_clicked = True
                domestic_radio_element = radio
                break
            except Exception:
                continue

        # Strategy 2: Search all radios
        if not radio_clicked:
            all_radios = driver.find_elements(By.XPATH, "//input[@type='radio']")
            logging.info(f"   Searching {len(all_radios)} radio buttons...")
            for r in all_radios:
                try:
                    label_text = r.find_element(By.XPATH, "./..").text.lower()
                except Exception:
                    label_text = ""
                if "create domestic customer" in label_text:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", r)
                    time.sleep(0.2)
                    reporter.log_step("Found_Domestic_Radio_S2", "Found radio (Strategy 2), about to click")
                    try:
                        r.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", r)
                    logging.info("Create Domestic Customer radio clicked (Strategy 2)")
                    reporter.log_step("Click_Domestic_Radio", "Clicked 'Create Domestic Customer' radio")
                    radio_clicked = True
                    domestic_radio_element = r
                    break

        if not radio_clicked:
            logging.error(" Could not click 'Create Domestic Customer' radio")
            reporter.log_error("Click_Domestic_Radio_Failed", "Could not find/click radio button")
            return False

        time.sleep(1)

        # Click Next button
        logging.info(" Looking for Next button...")
        next_clicked = False

        # Method 1: Find Next in same container as radio
        for text_elem in domestic_elements:
            try:
                card_container = text_elem.find_element(
                    By.XPATH,
                    "./ancestor::*[.//button[contains(@class,'slds-button') and contains(., 'Next')]][1]"
                )
                next_button = card_container.find_element(
                    By.XPATH,
                    ".//button[contains(@class,'slds-button_brand') and contains(., 'Next')]"
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(0.3)
                reporter.log_step("Found_Next_Button_M1", "Found Next button (Method 1), about to click")
                try:
                    next_button.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", next_button)
                logging.info(" Next button clicked (Method 1)")
                reporter.log_step("Click_Next_Button", "Clicked Next button")
                next_clicked = True
                break
            except Exception:
                continue

        # Method 2: From radio element
        if not next_clicked and domestic_radio_element:
            try:
                card_from_radio = domestic_radio_element.find_element(
                    By.XPATH,
                    "./ancestor::*[.//button[contains(@class,'slds-button_brand') and contains(., 'Next')]][1]"
                )
                next_in_card = card_from_radio.find_element(
                    By.XPATH,
                    ".//button[contains(@class,'slds-button_brand') and contains(., 'Next')]"
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_in_card)
                time.sleep(0.3)
                reporter.log_step("Found_Next_Button_M2", "Found Next button (Method 2), about to click")
                try:
                    next_in_card.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", next_in_card)
                logging.info(" Next button clicked (Method 2)")
                reporter.log_step("Click_Next_Button", "Clicked Next button")
                next_clicked = True
            except Exception:
                pass

        # Method 3: Find any Next button (avoid navigation)
        if not next_clicked:
            all_next = driver.find_elements(By.XPATH, "//button[contains(@class,'slds-button') and contains(., 'Next')]")
            logging.info(f"   Found {len(all_next)} Next buttons")
            for idx, btn in enumerate(all_next):
                try:
                    # Skip navigation buttons
                    try:
                        btn.find_element(By.XPATH, "./ancestor::*[contains(@class,'navigate') or contains(@class,'flow-runtime-navigate')]")
                        continue
                    except Exception:
                        pass

                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                    time.sleep(0.2)
                    reporter.log_step("Found_Next_Button_M3", f"Found Next button (Method 3, index {idx}), about to click")
                    try:
                        btn.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn)
                    logging.info(f" Next button clicked (Method 3, button {idx})")
                    reporter.log_step("Click_Next_Button", f"Clicked Next button (fallback {idx})")
                    next_clicked = True
                    break
                except Exception:
                    continue

        if not next_clicked:
            logging.error(" FAILED TO CLICK NEXT BUTTON")
            reporter.log_error("Click_Next_Failed", "Could not click Next button")
            return False

        logging.info(" Waiting for form page to load...")
        time.sleep(5)

        # Verify form loaded
        form_inputs = driver.find_elements(By.XPATH, "//input[@type='text' or @type='email' or @type='tel']")
        logging.info(f"   Found {len(form_inputs)} form inputs")
        
        if len(form_inputs) >= 3:
            logging.info(" FORM PAGE LOADED SUCCESSFULLY")
            reporter.log_step("Form_Page_Loaded", f"Form loaded with {len(form_inputs)} inputs")
            return True
        else:
            logging.error(" Form page did not load properly")
            reporter.log_error("Form_Load_Failed", "Form inputs not found")
            return False

    except Exception as e:
        logging.error(f" Navigation failed: {str(e)}")
        logging.error(traceback.format_exc())
        reporter.log_error("Navigate_Form_Exception", "Navigation failed", e)
        return False


# =========================
# PICKLIST HELPER FUNCTIONS
# =========================
def _norm(s: str) -> str:
    """Normalize string for comparison"""
    return " ".join(s.strip().lower().split())


def _options(sel):
    """Get all non-empty options from a select element"""
    return [o.text.strip() for o in Select(sel).options if o.text and o.text.strip()]


def _pick_by_visible_text(sel, value) -> bool:
    """Select option by visible text with fuzzy matching"""
    try:
        Select(sel).select_by_visible_text(value)
        return True
    except Exception:
        pass
    
    try:
        opts = _options(sel)
        for o in opts:
            if _norm(o) == _norm(value) or _norm(value) in _norm(o):
                Select(sel).select_by_visible_text(o)
                return True
    except Exception:
        pass
    
    return False


def _find_select_by_ids_or_label(driver, label_hint: str, preferred_ids=None):
    """Find select element by preferred IDs or label"""
    if preferred_ids:
        for pid in preferred_ids:
            try:
                el = driver.find_element(By.ID, pid)
                if el.tag_name.lower() == "select":
                    return el
            except Exception:
                pass
    
    try:
        labels = driver.find_elements(
            By.XPATH,
            f"//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{_norm(label_hint)}')]"
        )
        for lab in labels:
            for_attr = lab.get_attribute("for")
            if for_attr:
                try:
                    el = driver.find_element(By.ID, for_attr)
                    if el.tag_name.lower() == "select":
                        return el
                except Exception:
                    pass
            try:
                el = lab.find_element(By.XPATH, ".//following::select[1]")
                return el
            except Exception:
                pass
    except Exception:
        pass
    
    return None


def _find_all_selects(driver):
    """Find all select elements on the page"""
    try:
        return driver.find_elements(By.XPATH, "//select")
    except Exception:
        return []


def _guess_division_select(driver):
    """Find the Division select element"""
    el = _find_select_by_ids_or_label(driver, "Division", PREFERRED_IDS.get("Division"))
    if el:
        return el
    
    for sel in _find_all_selects(driver):
        opts = set(_options(sel))
        if any(v in opts for v in ALLOWED_DIVISIONS):
            return sel
    
    sels = _find_all_selects(driver)
    return sels[0] if sels else None


def _guess_sector_select(driver):
    """Find the Sector Type select element"""
    el = _find_select_by_ids_or_label(driver, "Sector Type", PREFERRED_IDS.get("Sector Type"))
    if el:
        return el
    
    for sel in _find_all_selects(driver):
        joined = _norm(" ".join(_options(sel)))
        keywords = ["education", "health", "retail", "office", "services", "nhs", "hotel", "agriculture", "leisure", "manufacturing"]
        if any(k in joined for k in keywords):
            return sel
    
    return None


def _guess_business_select(driver):
    """Find the Business Type select element"""
    el = _find_select_by_ids_or_label(driver, "Business Type", PREFERRED_IDS.get("Business Type"))
    if el:
        return el
    
    sels = _find_all_selects(driver)
    return sels[-1] if sels else None


# ==========================================================
# NEW/IMPROVED LWC PICKLIST HELPER
# ==========================================================
def select_lwc_combobox_option(driver, reporter: AutomationReporter, label: str, option: str, timeout: int = 10) -> bool:
    """
    Attempts to select an option from a Lightning Web Component (LWC) combobox.
    This is a robust function with multiple fallbacks.
    
    :param driver: The Selenium WebDriver instance.
    :param reporter: The AutomationReporter for logging.
    :param label: The visible text of the label (e.g., "Division", "Sector Type").
    :param option: The exact text of the option to select (e.g., "Homeowner").
    :param timeout: How long to wait for elements.
    :return: True if successful, False otherwise.
    """
    try:
        wait = WebDriverWait(driver, timeout)
        logging.info(f"  [LWC] Attempting combobox '{label}' -> '{option}'")
        
        # 1. Find the combobox input using its label.
        # Use 'contains' for flexibility (e.g., "Division" vs "Division *")
        label_xpath = f"//label[contains(normalize-space(.), '{label}')]"
        label_element = wait.until(EC.presence_of_element_located((By.XPATH, label_xpath)))
        
        # Find the 'for' attribute to get the input ID (most reliable)
        input_id = label_element.get_attribute("for")
        
        combobox_input = None
        if input_id:
            try:
                combobox_input = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
            except Exception:
                logging.warning(f"  [LWC] Found label for '{label}' but ID '{input_id}' not clickable. Trying relative search.")
        
        if not combobox_input:
            # If no 'for' or ID failed, find the input relative to the label
            combobox_input = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"{label_xpath}/following::input[@role='combobox'][1]")
            ))

        if not combobox_input:
             raise Exception("Combobox input not found after all attempts")

        # 2. Click the combobox to open the options
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", combobox_input)
        time.sleep(0.5) # Wait for scroll
        driver.execute_script("arguments[0].click();", combobox_input) # JS click is more reliable
        
        logging.info(f"  [LWC] Clicked combobox for '{label}'")
        reporter.log_step(f"LWC_Click_Combobox_{label}", f"Clicked '{label}' dropdown")
        
        # 3. Wait for the listbox itself to appear
        listbox_xpath = "//div[contains(@class, 'slds-listbox') and @role='listbox']"
        wait.until(EC.visibility_of_element_located((By.XPATH, listbox_xpath)))
        time.sleep(0.2) # Wait for options to populate

        # 4. Find and click the desired option
        # STRATEGY 1 (Best): Find by span[@title='...']
        option_element = None
        try:
            option_xpath_title = f"//span[@title='{option}']/ancestor::div[contains(@class, 'slds-listbox__option') and @role='option'] | //span[@title='{option}']/ancestor::lightning-base-combobox-item"
            option_element = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath_title)))
            
            option_text = option_element.text
            logging.info(f"  [LWC] Found option by title: '{option_text}'. Clicking...")
            reporter.log_step(f"LWC_Found_Option_Title_{label}", f"Found '{option_text}', about to click")
        
        except Exception:
            # STRATEGY 2 (Fallback): Find by visible text
            logging.warning(f"  [LWC] Could not find option by title '{option}'. Trying by visible text...")
            option_xpath_text = f"//div[contains(@class, 'slds-listbox__option') and @role='option'][.//*[normalize-space(.)='{option}']]"
            option_element = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath_text)))
            
            option_text = option_element.text
            logging.info(f"  [LWC] Found option by text: '{option_text}'. Clicking...")
            reporter.log_step(f"LWC_Found_Option_Text_{label}", f"Found '{option_text}', about to click")

        # Click the found element
        driver.execute_script("arguments[0].click();", option_element)
        
        # 5. Wait for the dropdown to close (option is no longer visible)
        wait.until(EC.invisibility_of_element_located((By.XPATH, listbox_xpath)))
        reporter.log_step(f"LWC_Select_Option_{label}", f"Selected '{label}': '{option}'")
        logging.info(f"  [LWC] Successfully selected '{label}': '{option}'")
        return True
        
    except Exception as e:
        logging.error(f"  [LWC] All attempts FAILED for '{label}': '{option}'. Error: {e}")
        reporter.log_error(f"LWC_Select_Failed_{label}", f"Failed to select '{option}'", e)
        return False


# ==========================================================
# NEW/IMPROVED DIVISION/SECTOR/BUSINESS FUNCTION
# ==========================================================
def select_division_sector_business(driver, reporter: AutomationReporter, preferred_division: str):
    """
    Select Division, Sector Type, and Business Type.
    This now attempts LWC (combobox) selection first, then falls back
    to the original <select> element logic.
    """
    
    # --- 1. CHOOSE DIVISION ---
    chosen_div = None
    logging.info(f" Attempting to select Division: '{preferred_division}'")
    
    # Try LWC method
    if select_lwc_combobox_option(driver, reporter, "Division", preferred_division):
        chosen_div = preferred_division
    else:
        # Fallback to <select> method
        logging.info("  LWC method failed for Division, falling back to <select> method...")
        div_sel = _guess_division_select(driver) # Uses old logic
        if not div_sel:
            logging.warning("  No Division select found (LWC and <select> failed)")
            reporter.log_error("Select_Division_Failed", "Could not find Division dropdown (LWC or <select>)")
            return False
        
        div_opts = _options(div_sel)
        for candidate in [preferred_division] + ALLOWED_DIVISIONS:
            if any(_norm(o) == _norm(candidate) for o in div_opts):
                reporter.log_step("Found_Division_Select", f"Found <select> Division, about to select '{candidate}'")
                if _pick_by_visible_text(div_sel, candidate): # Uses old logic
                    chosen_div = candidate
                    logging.info(f" Division selected (select): {candidate}")
                    reporter.log_step("Select_Division", f"Division: {candidate}")
                    break
        if not chosen_div:
             logging.warning("  No allowed Division found in <select> dropdown")
             reporter.log_error("Select_Division_Failed", "No allowed Division option found in <select>")
             return False

    if not chosen_div:
        logging.error("  Failed to select Division by any method.")
        return False
    
    time.sleep(1.5) # Wait for dependent picklists to update

    # --- 2. CHOOSE SECTOR TYPE ---
    chosen_sector = None
    sector_candidates = list(DIVISION_TREE[chosen_div].keys())
    if not sector_candidates:
        logging.warning(f"  No Sector candidates found for Division '{chosen_div}'")
        return False # Cannot proceed
    
    # We'll just pick the first candidate for simplicity
    sector_to_select = sector_candidates[0]
    logging.info(f" Attempting to select Sector: '{sector_to_select}'")

    # Try LWC method
    if select_lwc_combobox_option(driver, reporter, "Sector Type", sector_to_select):
        chosen_sector = sector_to_select
    else:
        # Fallback to <select> method
        logging.info("  LWC method failed for Sector, falling back to <select> method...")
        sector_sel = _guess_sector_select(driver)
        if not sector_sel:
            logging.warning("  No Sector Type select found (LWC and <select> failed)")
            reporter.log_error("Select_Sector_Failed", "Could not find Sector Type select")
            return False
        
        sector_opts = _options(sector_sel)
        for s in sector_candidates:
            if any(_norm(o) == _norm(s) or _norm(s) in _norm(o) for o in sector_opts):
                reporter.log_step("Found_Sector_Select", f"Found <select> Sector, about to select '{s}'")
                if _pick_by_visible_text(sector_sel, s):
                    chosen_sector = s
                    logging.info(f" Sector Type selected (select): {s}")
                    reporter.log_step("Select_Sector_Type", f"Sector: {s}")
                    break
        
        if not chosen_sector and sector_opts:
            # Fallback to first available option
            reporter.log_step("Found_Sector_Select_Fallback", f"Using fallback, about to select '{sector_opts[0]}'")
            _pick_by_visible_text(sector_sel, sector_opts[0])
            chosen_sector = sector_opts[0]
            logging.info(f" Sector Type selected (select fallback): {chosen_sector}")
            reporter.log_step("Select_Sector_Type", f"Sector (fallback): {chosen_sector}")

    if not chosen_sector:
        logging.error("  Failed to select Sector by any method.")
        reporter.log_error("Select_Sector_Failed", "Sector options empty or not found")
        return False

    time.sleep(1.5) # Wait for dependent picklists to update

    # --- 3. CHOOSE BUSINESS TYPE ---
    chosen_biz = None
    # Use .get() for safety in case chosen_sector isn't a perfect match
    biz_candidates = DIVISION_TREE[chosen_div].get(chosen_sector, [])
    if not biz_candidates:
        logging.warning(f"  No Business candidates found for Sector '{chosen_sector}'")
        # Don't fail, some sectors might not have a business type
        reporter.log_step("Select_Business_Type", "No business types to select")
        return True # Not a failure
    
    # We'll just pick the first candidate
    biz_to_select = biz_candidates[0]
    logging.info(f" Attempting to select Business Type: '{biz_to_select}'")

    # Try LWC method
    if select_lwc_combobox_option(driver, reporter, "Business Type", biz_to_select):
        chosen_biz = biz_to_select
    else:
         # Fallback to <select> method
        logging.info("  LWC method failed for Business, falling back to <select> method...")
        biz_sel = _guess_business_select(driver)
        if not biz_sel:
            logging.warning("  No Business Type select found (LWC and <select> failed)")
            reporter.log_error("Select_Business_Failed", "Could not find Business Type select")
            return False
        
        biz_opts = _options(biz_sel)
        for b in biz_candidates:
            if any(_norm(o) == _norm(b) or _norm(b) in _norm(o) for o in biz_opts):
                reporter.log_step("Found_Business_Select", f"Found <select> Business, about to select '{b}'")
                if _pick_by_visible_text(biz_sel, b):
                    chosen_biz = b
                    logging.info(f" Business Type selected (select): {b}")
                    reporter.log_step("Select_Business_Type", f"Business: {b}")
                    break
        
        if not chosen_biz and biz_opts:
            # Fallback to first available option
            reporter.log_step("Found_Business_Select_Fallback", f"Using fallback, about to select '{biz_opts[0]}'")
            _pick_by_visible_text(biz_sel, biz_opts[0])
            chosen_biz = biz_opts[0]
            logging.info(f" Business Type selected (select fallback): {chosen_biz}")
            reporter.log_step("Select_Business_Type", f"Business (fallback): {chosen_biz}")
    
    if not chosen_biz:
        logging.warning(f"  No Business Type selected. This may be okay. Fallback options: {biz_opts}")
        # Don't fail here, as it might not be required.
        # If it *is* required, the form submission will fail, which is correct.
        reporter.log_error("Select_Business_Failed", "Business Type options empty or not found")


    return True


# =========================
# ADDRESS HELPER FUNCTIONS
# =========================
def _normalize_uk_postcode(raw: str) -> str:
    """Uppercase and ensure single space before inward code (e.g., SW1A 1AA)"""
    s = re.sub(r"\s+", "", str(raw).upper())
    if len(s) >= 5:
        return s[:-3] + " " + s[-3:]
    return s


def _compose_full_address(building_number: str, address1: str, city: str) -> str:
    """Compose full address from components"""
    parts = []
    if building_number:
        parts.append(str(building_number).strip())
    if address1:
        parts.append(str(address1).strip())
    if city:
        parts.append(str(city).strip())
    return ", ".join([p for p in parts if p])


def fill_address_search_with_full_address(driver, reporter: AutomationReporter, full_address: str, building_number: str, timeout: int = 10) -> bool:
    """Type full address, wait for suggestions, select best match"""
    if not full_address:
        return False

    search_xpaths = [
        "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'search for address')]",
        "//label[contains(., 'Address')]/following::input[1]",
        "//input[@role='combobox' and (contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'address') or contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'address'))]",
        "(//label[contains(.,'Address')]/ancestor::*[self::div or self::lightning-input][1]//input[not(@type='hidden')])[1]",
    ]

    addr_input = None
    for xp in search_xpaths:
        try:
            addr_input = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.XPATH, xp))
            )
            break
        except Exception:
            continue

    if not addr_input:
        logging.error(" Address search input not found")
        reporter.log_error("Address_Search_Input_Failed", "Could not find Address search input")
        return False

    try:
        reporter.log_step("Found_Address_Input", "Found address search input, about to click")
        addr_input.click()
    except Exception:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", addr_input)
        driver.execute_script("arguments[0].click();", addr_input)

    # Clear existing text
    try:
        addr_input.send_keys(Keys.CONTROL, "a")
    except Exception:
        try:
            addr_input.send_keys(Keys.COMMAND, "a")
        except Exception:
            pass
    
    addr_input.send_keys(Keys.DELETE)
    reporter.log_step("Address_Input_Cleared", "Cleared address input, about to type")
    
    addr_input.send_keys(full_address)
    reporter.log_step("Address_Search_Typed", f"Typed address: {full_address}")

    # Trigger suggestions
    try:
        addr_input.send_keys(" ")
        addr_input.send_keys(Keys.BACK_SPACE)
    except Exception:
        pass

    # Wait for suggestions
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'slds-listbox__option')]"))
        )
        reporter.log_step("Address_Suggestions_Found", "Address suggestions appeared")
    except TimeoutException:
        logging.error(" No address suggestions appeared")
        reporter.log_error("Address_Suggestion_Failed", "No suggestions appeared")
        return False

    suggestions = driver.find_elements(By.XPATH, "//div[contains(@class,'slds-listbox__option')]")
    if not suggestions:
        logging.error(" Address suggestions list empty")
        reporter.log_error("Address_Suggestion_Failed", "Suggestions list empty")
        return False

    # Select best match
    chosen = None
    bn = str(building_number).strip() if building_number else ""
    
    # Try to match building number
    if bn:
        for s in suggestions:
            try:
                t = s.text.strip()
                if re.search(rf"\b{re.escape(bn)}\b", t):
                    chosen = s
                    break
            except Exception:
                continue
    
    # Try to match start of address
    if not chosen:
        for s in suggestions:
            try:
                if s.text and s.text.strip().lower().startswith(str(full_address).split(",")[0].strip().lower()):
                    chosen = s
                    break
            except Exception:
                continue
    
    # Fallback to first suggestion
    if not chosen:
        chosen = suggestions[0]

    chosen_text = chosen.text
    reporter.log_step("Found_Address_Suggestion", f"Found suggestion '{chosen_text}', about to click")
    driver.execute_script("arguments[0].click();", chosen)
    logging.info(f" Address suggestion selected: {chosen_text}")
    reporter.log_step("Address_Suggestion_Clicked", f"Selected: {chosen_text}")
    time.sleep(0.7)
    return True


# =========================
# FORM FILLING FUNCTION
# =========================
def fill_form(driver, reporter: AutomationReporter, row_data, row_number):
    """Fill the form with data from a single row"""
    try:
        wait = WebDriverWait(driver, 20)
        logging.info(f" Filling form for row {row_number}")
        time.sleep(2)

        def fill_field(field_name, field_value, selectors, step_name_prefix, reporter: AutomationReporter):
            """Helper function to fill a single field"""
            if pd.isna(field_value) or str(field_value).strip() == "":
                logging.warning(f"   {field_name} is empty, skipping")
                return False

            for selector in selectors:
                try:
                    field = driver.find_element(By.XPATH, selector)
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", field)
                    time.sleep(0.2)
                    
                    reporter.log_step(f"Attempt_{step_name_prefix}", f"About to fill {field_name}")

                    try:
                        field.clear()
                        field.send_keys(str(field_value))
                    except Exception:
                        driver.execute_script(
                            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                            field, str(field_value)
                        )
                    
                    logging.info(f" {field_name}: {field_value}")
                    reporter.log_step(f"Filled_{step_name_prefix}", f"{field_name}: {field_value}")
                    return True
                except Exception:
                    continue

            logging.error(f" Could not fill {field_name}")
            reporter.log_error(f"Fill_{step_name_prefix}_Failed", f"Could not find/fill {field_name}")
            return False

        # FIRST NAME
        fill_field("First Name", get_row_value(row_data, ['FirstName', 'First Name', 'FName']), [
            "//label[contains(., 'First Name')]/following::input[1]",
            "//label[contains(., 'First')]/following::input[1]",
            "//input[contains(@placeholder, 'First')]",
            "//input[contains(@name, 'FirstName')]",
            "(//input[@type='text'])[1]"
        ], "FirstName", reporter)

        # LAST NAME
        fill_field("Last Name", get_row_value(row_data, ['LastName', 'Last Name', 'Surname', 'LName']), [
            "//label[contains(., 'Last Name')]/following::input[1]",
            "//label[contains(., 'Last')]/following::input[1]",
            "//input[contains(@placeholder, 'Last')]",
            "//input[contains(@name, 'LastName')]",
            "(//input[@type='text'])[2]"
        ], "LastName", reporter)

        # PHONE
        fill_field("Phone", get_row_value(row_data, ['Phone', 'Telephone', 'Mobile']), [
            "//label[contains(., 'Phone')]/following::input[1]",
            "//input[@type='tel']",
            "//input[contains(@name, 'Phone')]",
            "//input[contains(@placeholder, 'Phone')]"
        ], "Phone", reporter)

        # EMAIL
        fill_field("Email", get_row_value(row_data, ['Email', 'E-mail']), [
            "//label[contains(., 'Email')]/following::input[1]",
            "//input[@type='email']",
            "//input[contains(@name, 'Email')]",
            "//input[contains(@placeholder, 'Email')]"
        ], "Email", reporter)

        # REFERRED - Select "No"
        try:
            radios = driver.find_elements(By.XPATH, "//input[@type='radio']")
            for r in radios:
                try:
                    label_text = ""
                    try:
                        label_text = r.find_element(By.XPATH, "./ancestor::span[1]").text.lower()
                    except Exception:
                         label_text = r.find_element(By.XPATH, "./..").text.lower()

                    if "no" in label_text:
                        # Check if parent text mentions "referred"
                        parent_text = r.find_element(By.XPATH, "./ancestor::fieldset[1] | ./ancestor::div[.//legend][1]").text.lower()
                        if "referred" in parent_text:
                            reporter.log_step("Attempt_Select_Referred_No", f"Found 'No' radio for 'Referred', about to click")
                            driver.execute_script("arguments[0].click();", r)
                            logging.info("Referred: No")
                            reporter.log_step("Select_Referred_No", "Selected 'Referred: No'")
                            break
                except Exception:
                    continue
        except Exception as e:
            logging.warning(f"   Could not select 'Referred: No': {e}")
            reporter.log_error("Select_Referred_No_Failed", "Could not select 'Referred: No'", e)


        # ==========================================================
        # CHANGED: DIVISION / SECTOR / BUSINESS
        # ==========================================================
        # NEW: Pick a random division from the allowed list
        random_division = random.choice(ALLOWED_DIVISIONS)
        logging.info(f"  Attempting to select random division: {random_division}")
        
        if not select_division_sector_business(driver, reporter, preferred_division=random_division):
            logging.warning("  Division/Sector/Business selection incomplete. This may cause validation errors.")
            # Do not return false, allow the script to try and submit anyway
            # The validation error check will catch it if it's a problem.


        # BUILDING NUMBER
        building_num_val = get_row_value(
            row_data,
            ['BuildingNumber', 'Building Number', 'HouseNumber', 'House Number', 'House name/number', 'House No']
        )
        fill_field("Building Number", building_num_val, [
            "//lightning-input[.//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'building') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'number')]]//input",
            "//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'building') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'number')]/following::input[1]",
            "//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'building')]",
            "//input[contains(@name, 'Building') or contains(@name, 'building')]",
        ], "BuildingNumber", reporter)

        # ADDRESS - Use full address search
        address1_val = get_row_value(row_data, ['AddressLine1', 'Address1', 'Street', 'Address Line 1'])
        city_val = get_row_value(row_data, ['City', 'Town'])
        postcode_val = get_row_value(row_data, ['Postcode', 'PostalCode', 'ZIP'])

        full_address = _compose_full_address(building_num_val, address1_val, city_val)
        
        if full_address:
            try:
                fill_address_search_with_full_address(
                    driver=driver,
                    reporter=reporter,
                    full_address=full_address,
                    building_number=building_num_val,
                    timeout=10
                )
            except Exception as e:
                logging.exception(f" Address search error: {e}")
                reporter.log_error("Address_Search_Exception", "Address search failed", e)
        else:
            reporter.log_error("Address_Search_Skipped", "Full address empty")

        # POSTAL CODE - Explicit fill
        if postcode_val:
            normalized_pc = _normalize_uk_postcode(postcode_val)
            try:
                pc_input = None
                for xp in [
                    "//label[contains(., 'Postal Code') or contains(., 'Postcode')]/following::input[1]",
                    "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'postcode') or contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'postcode')]",
                ]:
                    try:
                        pc_input = WebDriverWait(driver, 6).until(EC.visibility_of_element_located((By.XPATH, xp)))
                        break
                    except Exception:
                        continue
                
                if pc_input:
                    reporter.log_step("Attempt_Fill_Postal_Code", "Found Postal Code input, about to fill")
                    try:
                        pc_input.click()
                    except Exception:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pc_input)
                        driver.execute_script("arguments[0].click();", pc_input)
                    
                    try:
                        pc_input.send_keys(Keys.CONTROL, "a")
                    except Exception:
                        try:
                            pc_input.send_keys(Keys.COMMAND, "a")
                        except Exception:
                            pass
                    
                    pc_input.send_keys(Keys.DELETE)
                    pc_input.send_keys(normalized_pc)
                    logging.info(f" Postal Code: {normalized_pc}")
                    reporter.log_step("Fill_Postal_Code", f"Postal Code: {normalized_pc}")
            except Exception as e:
                logging.warning(f"   Could not fill Postal Code: {e}")

        # STREET / CITY - Fallbacks
        fill_field("Street", address1_val, [
            "//label[contains(., 'Street')]/following::input[1]",
            "//input[contains(@name, 'Street')]"
        ], "Street", reporter)
        
        fill_field("City", city_val, [
            "//label[contains(., 'City')]/following::input[1]",
            "//input[contains(@name, 'City')]"
        ], "City", reporter)

        time.sleep(1)

        # SUBMIT / NEXT BUTTON
        logging.info(" Looking for Submit/Next button...")
        submit_clicked = False

        buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'slds-button_brand')]")
        for btn in buttons:
            try:
                txt = btn.text.strip().lower()
                if any(k in txt for k in ["next", "submit", "save", "create"]):
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                    time.sleep(0.2)
                    reporter.log_step("Attempt_Form_Submit", f"Found button '{btn.text.strip()}', about to click")
                    try:
                        btn.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn)
                    logging.info(f" Clicked button: '{btn.text.strip()}'")
                    reporter.log_step("Form_Submit", f"Clicked: '{btn.text.strip()}'")
                    submit_clicked = True
                    break
            except Exception:
                continue

        if not submit_clicked:
            buttons = driver.find_elements(By.XPATH, "//button[contains(., 'Next') or contains(., 'Submit') or contains(., 'Save') or contains(., 'Create')]")
            for btn in buttons:
                try:
                    reporter.log_step("Attempt_Form_Submit_FB", f"Found fallback button '{btn.text.strip()}', about to click")
                    driver.execute_script("arguments[0].click();", btn)
                    logging.info(f" Clicked button (fallback): '{btn.text.strip()}'")
                    reporter.log_step("Form_Submit", f"Clicked (fallback): '{btn.text.strip()}'")
                    submit_clicked = True
                    break
                except Exception:
                    continue

        if not submit_clicked:
            logging.error(" Could not click Submit/Next button")
            reporter.log_error("Form_Submit_Failed", "Could not click Submit/Next")
            return False

        time.sleep(3)

        # CHECK FOR VALIDATION ERRORS
        error_elements = driver.find_elements(By.XPATH, "//*[contains(., 'Complete this field') or contains(., 'required')]")
        if error_elements:
            logging.error(f" Validation errors for row {row_number}")
            error_text = ""
            for err in error_elements[:5]:
                try:
                    e_text = err.text
                    logging.error(f"   Error: {e_text}")
                    error_text += f"{e_text}\n"
                except Exception:
                    pass
            reporter.log_error("Validation_Error", f"Validation errors (row {row_number}):\n{error_text}")
            return False
        else:
            logging.info(f" Row {row_number} submitted successfully!")
            reporter.log_step("Form_Submit_Success", f"Row {row_number} SUCCESS")
            return True

    except Exception as e:
        logging.error(f" Error filling form for row {row_number}: {str(e)}")
        logging.error(traceback.format_exc())
        reporter.log_error(f"Form_Fill_Exception_Row_{row_number}", "Form fill failed", e)
        return False


# =========================
# MAIN FUNCTION
# =========================
def main():
    driver = None
    reporter = None
    successful_rows = []
    failed_rows = []

    try:
        # LOAD CSV
        logging.info("=" * 70)
        logging.info(" LOADING DATA")
        logging.info("=" * 70)
        logging.info(f"Loading CSV: {CSV_FILE}")
        
        df = pd.read_csv(CSV_FILE)
        logging.info(f" Loaded {len(df)} rows")
        logging.info(f"Columns: {list(df.columns)}")

        # INITIALIZE DRIVER
        driver = initialize_driver()
        
        # INITIALIZE REPORTER
        reporter = AutomationReporter(driver, base_run_name="Salesforce_Automation_Test")
        reporter.log_step("Init_Driver", "ChromeDriver initialized")

        # LOGIN TO SALESFORCE
        logging.info("=" * 70)
        logging.info(" LOGGING IN TO SALESFORCE")
        logging.info("=" * 70)
        
        if not login_to_salesforce(driver, reporter):
            logging.error(" Login failed - aborting")
            return

        logging.info(" Login successful!")

        # PROCESS ROWS
        logging.info("=" * 70)
        logging.info(f" PROCESSING {len(df)} ROWS")
        logging.info("=" * 70)

        for index, row in df.iterrows():
            row_number = index + 1
            logging.info("\n" + "=" * 70)
            logging.info(f"ROW {row_number} of {len(df)}")
            logging.info("=" * 70)
            reporter.log_step(f"Start_Row_{row_number}", f"Processing row {row_number}/{len(df)}")

            try:
                if navigate_to_form(driver, reporter):
                    if fill_form(driver, reporter, row, row_number):
                        successful_rows.append(row_number)
                        logging.info(f" ROW {row_number} SUCCESS")
                        reporter.log_step(f"Row_{row_number}_Success", f"Row {row_number} completed")
                    else:
                        failed_rows.append(row_number)
                        logging.error(f"ROW {row_number} FAILED")
                else:
                    failed_rows.append(row_number)
                    logging.error(f" ROW {row_number} NAV FAILED")

                time.sleep(2)

            except Exception as e:
                failed_rows.append(row_number)
                logging.error(f" ROW {row_number} ERROR: {str(e)}")
                logging.error(traceback.format_exc())
                reporter.log_error(f"Row_{row_number}_Exception", "Unhandled exception", e)

        # FINAL SUMMARY
        logging.info("\n" + "=" * 70)
        logging.info(" FINAL SUMMARY")
        logging.info("=" * 70)
        logging.info(f"Total Rows:     {len(df)}")
        logging.info(f" Successful:   {len(successful_rows)}")
        logging.info(f" Failed:       {len(failed_rows)}")
        
        if failed_rows:
            logging.info(f"Failed rows: {failed_rows}")
        
        reporter.log_step("Run_Finished", f"Automation complete. Success: {len(successful_rows)}, Failed: {len(failed_rows)}")

    except Exception as e:
        logging.error(f" CRITICAL ERROR: {str(e)}")
        logging.error(traceback.format_exc())
        if reporter:
            reporter.log_error("Main_Exception", "Critical error in main function", e)
            
    finally:
        # CREATE ARTIFACTS ZIP
        if reporter:
            logging.info("\n" + "=" * 70)
            logging.info(" FINALIZING AND ZIPPING ARTIFACTS")
            logging.info("=" * 70)
            
            # This new function saves the report AND creates the zip file
            reporter.create_artifacts_zip()
            
        if driver:
            logging.info(" Closing WebDriver")
            driver.quit()
        
        logging.info("=" * 70)
        logging.info(" SCRIPT FINISHED")
        logging.info("=" * 70)


if __name__ == "__main__":
    main()
