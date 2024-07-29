
"""
LoginTest.py
Program: DDTF main executing file
"""

# Import necessary modules and classes for the script
from Locators.Homepage import OrangeHRMLocators
from utilities.excel_functions import Excel_functions
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# Retrieve the Excel file path and sheet number from the locator class
excel_file=OrangeHRMLocators().excel_file
sheet_number=OrangeHRMLocators().sheet_number

# Initialize the Excel_functions class with the file and sheet number
excel=Excel_functions(excel_file,sheet_number)
# Set up the Chrome WebDriver using ChromeDriverManager for automatic driver management
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()))
# Open the URL specified in the OrangeHRMLocators class
driver.get(OrangeHRMLocators().url)
# Create a WebDriverWait instance for waiting up to 10 seconds for elements to become visible or clickable
wait=WebDriverWait(driver,10)
# Get the total number of rows in the Excel sheet
row=excel.row_count()

# Loop through each row of data starting from row 2 to the last row
for row in range(2,row+1):
    # Read the username and password from the Excel file
    username=excel.read_data(row,6)
    password=excel.read_data(row,7)

    # Wait until the username field is visible and then enter the username
    username_locator=wait.until(EC.visibility_of_element_located((By.NAME, OrangeHRMLocators().username)))
    username_locator.send_keys(username)

    # Wait until the password field is visible and then enter the password
    password_locator=wait.until(EC.visibility_of_element_located((By.NAME,OrangeHRMLocators().password)))
    password_locator.send_keys(password)

    # Wait until the submit button is clickable and then click it
    submit_button=wait.until(EC.element_to_be_clickable((By.XPATH,OrangeHRMLocators().submit_button)))
    submit_button.click()

    # Check if the current URL matches the dashboard URL indicating a successful login
    if OrangeHRMLocators().dashboard_url==driver.current_url:
        print(f"SUCCESS : Login with username {username} and password {password} ")

        # Write a success status to the Excel file
        excel.write_data(row,8,OrangeHRMLocators().pass_data)
        # Navigate back to the login page
        driver.back()

    # Check if the current URL matches the login URL indicating a failed login
    elif OrangeHRMLocators().url == driver.current_url:
        print(f"ERROR : Login unsuccessful with username {username} and password {password}")
        # Write a failure status to the Excel file
        excel.write_data(row,8,OrangeHRMLocators().fail_data)
        # Refresh the page to retry
        driver.refresh()
# Close the browser after processing all rows
driver.quit()

