# Automated-Testing-Tool-Salesforces-customer-filling-

Salesforce Automation Script 
Overview
This Python script automates the process of creating  customer records in Salesforce. It reads customer data from a CSV file and fills out Salesforce forms automatically using Selenium WebDriver.
Features

Automated Login: Logs into Salesforce with credentials from environment variables
Batch Processing: Processes multiple customer records from CSV file
Smart Form Filling: Automatically fills customer information including:

Personal details (Name, Phone, Email)
Address information with intelligent address search
Division, Sector, and Business Type selection


Comprehensive Logging: Takes screenshots at every step and maintains detailed logs
Error Handling: Detects validation errors and continues with remaining records
Artifact Generation: Creates a ZIP file with all screenshots, logs, and JSON reports

Prerequisites

Python 3.8 or higher
Google Chrome browser installed
Salesforce sandbox account access
CSV file with customer data

Installation
1. Clone or Download the Script
Save the Python script to your local machine.
2. Install Required Packages
Create a requirements.txt file with the following content:
txtselenium==4.15.2
webdriver-manager==4.0.1
pandas==2.1.3
python-dotenv==1.0.0
Pillow==10.1.0
Install the packages:
bashpip install -r requirements.txt
3. Set Up Environment Variables
Create a .env file in the same directory as your script:
envSF_USERNAME=kunguma.balaji@aspect.co.uk.staging
SF_PASSWORD= .env 
DEFAULT_DIVISION=Homeowner
Important: Replace your_actual_password_here with your actual Salesforce password.
4. Prepare Your CSV Data File
Create or place your CSV file at: C:\Users\User\Downloads\salesforce_synthetic_data.csv
Required CSV Columns (the script supports multiple column name variations):

First Name: FirstName, First Name, or FName
Last Name: LastName, Last Name, Surname, or LName
Phone: Phone, Telephone, or Mobile
Email: Email or E-mail
Building Number: BuildingNumber, Building Number, HouseNumber, or House No
Address Line 1: AddressLine1, Address1, Street, or Address Line 1
City: City or Town
Postcode: Postcode, PostalCode, or ZIP

Example CSV format:
csvFirstName,LastName,Phone,Email,BuildingNumber,AddressLine1,City,Postcode
John,Smith,07700900123,john.smith@example.com,123,High Street,London,SW1A 1AA
Jane,Doe,07700900456,jane.doe@example.com,45,Baker Street,Manchester,M1 1AA
Configuration
Salesforce Configuration
Update these variables in the script if needed:
pythonLOGIN_URL = "https://test.salesforce.com/"
HOME_URL = "https://chumley--staging.sandbox.lightning.force.com/lightning/page/home"
CSV_FILE = r"C:\Users\User\Downloads\salesforce_synthetic_data.csv"
Division Options
The script randomly selects from these divisions:

Homeowner
Insurance
Corporate
Government

Each division has specific Sector Types and Business Types that are automatically selected.
Usage
Run the Script
bashpython your_script_name.py
```

### Manual Verification Code
If Salesforce requires a verification code:
1. The script will pause for 40 seconds
2. Enter the verification code in the browser window
3. The script will automatically continue

### Monitoring Progress
The script provides real-time console output showing:
- Current row being processed
- Field values being entered
- Success/failure status for each record
- Final summary with success/failure counts

## Output

### Artifacts Directory
The script creates an `artifacts` folder containing:
```
artifacts/
└── Automation_Run_YYYY-MM-DD_HH-MM-SS/
    ├── 001_navigate_login.png
    ├── 002_found_username_field.png
    ├── 003_enter_username.png
    ├── ... (screenshot for every step)
    ├── automation_run.log
    └── report.json
ZIP File
A compressed ZIP file is created: Automation_Run_YYYY-MM-DD_HH-MM-SS_artifacts.zip
Log Files

Console Log: automation_main.log - Main script execution log
Run-specific Log: automation_run.log - Detailed step-by-step log inside artifacts folder
JSON Report: report.json - Structured report with timestamps and screenshots

Troubleshooting
Common Issues
1. ChromeDriver Error

The script automatically downloads the correct ChromeDriver
Ensure Google Chrome is installed and up-to-date

2. Login Fails

Verify credentials in .env file
Check if verification code is required (script will pause)
Ensure Salesforce sandbox URL is correct

3. CSV File Not Found

Check the file path in CSV_FILE variable
Use raw string notation: r"C:\path\to\file.csv"

4. Form Fields Not Found

Salesforce UI may have changed
Check the logs and screenshots in artifacts folder
Update XPath selectors if necessary

5. Address Search Not Working

Ensure full address is properly formatted
Check if address search dropdown appears
Verify postcode format (UK format: "SW1A 1AA")

6. Division/Sector/Business Selection Fails

The script tries both LWC (Lightning Web Component) and traditional <select> methods
Check screenshots to see which elements are present
Ensure Division tree configuration matches Salesforce setup

Script Workflow

Initialization

Load CSV data
Initialize Chrome WebDriver
Create AutomationReporter for logging


Login

Navigate to Salesforce login page
Enter credentials
Handle verification code if required


For Each Row in CSV

Navigate to form page
Select "Create Domestic Customer" option
Fill personal information fields
Select Division, Sector, and Business Type (randomly)
Search and select address
Submit form
Check for validation errors


Finalization

Generate summary report
Create ZIP file with all artifacts
Close browser



Advanced Configuration
Customizing Division Selection
To use a specific division instead of random selection, modify this line in fill_form():
python# Change from:
random_division = random.choice(ALLOWED_DIVISIONS)

To
random_division = "Corporate"  # or your preferred division
Adding Custom Field Mappings
To add support for additional CSV column names, update the fill_field() calls:
pythonfill_field("First Name", get_row_value(row_data, ['FirstName', 'First Name', 'FName', 'YourCustomColumnName']), ...)
Adjusting Timeouts
Modify wait times if Salesforce is slow:
pythonwait = WebDriverWait(driver, 30)  # Increase from 30 to 60 seconds
Security Notes

Never commit .env file to version control
Add .env to your .gitignore file
Store passwords securely
Use Salesforce sandbox environment for testing
