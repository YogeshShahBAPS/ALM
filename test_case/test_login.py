from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from page_objects.login import LoginPage
from utilities.excel import get_matched_test_cases
from datetime import datetime
import pandas as pd
import os
import sys

# File paths
test_case_file = r"E:\Priojecttree\Automation Testing\Git\ALM\utilities\Test_Case.xlsx"
data_file = r"E:\Priojecttree\Automation Testing\Git\ALM\utilities\ALM_Data.xlsx"
screenshot_dir = r"E:\Priojecttree\Automation Testing\Git\ALM\ScreenShots"
output_dir = r"E:\Priojecttree\Automation Testing\Git\ALM\Excel Report"

# Ensure directories exist
os.makedirs(screenshot_dir, exist_ok=True)
os.makedirs(output_dir, exist_ok=True)

# Call the utility function to get matched test cases
matched_data = get_matched_test_cases(test_case_file, data_file)

# Check if matched data is empty
if matched_data.empty:
    print("No matched test cases to execute.")
    sys.exit()

# Initialize the WebDriver
driver = webdriver.Chrome()
driver.maximize_window()

# Initialize the page object
login_page = LoginPage(driver)

# Create a WebDriverWait instance
wait = WebDriverWait(driver, 10)

# Iterate through matched test cases
for index, row in matched_data.iterrows():
    tc_number = row['MTC']
    user_name = row['username']
    password = row['password']
    expected_text = row['Expected Text']

    result = "FAIL"  # Default result

    try:
        # Navigate to the login page
        driver.get("https://qa.baps.dev/alm/signIn")

        # Perform login using page object methods
        login_page.click_login_with_baps_sso()
        login_page.enter_username(user_name)
        login_page.enter_password(password)
        login_page.click_sign_in()

        # Verify dashboard text
        dashboard_xpath = "//*[@id='main_content']/app-dashboard/div[1]/div/h4"
        actual_text = wait.until(EC.visibility_of_element_located((By.XPATH, dashboard_xpath))).text

        if actual_text.strip() == expected_text.strip():
            result = "PASS"

    except Exception as e:
        screenshot_path = os.path.join(screenshot_dir, f"{tc_number}.png")
        driver.save_screenshot(screenshot_path)
        print(f"Error during test case {tc_number}: {e}")
        result = "FAIL"

    # Update the result in the matched data
    matched_data.at[index, 'TestCase_Result'] = result

# Close the WebDriver
driver.quit()

# Save the results to an Excel file
current_datetime = datetime.now().strftime("%d-%m_%H-%M-%S")
output_file_path = os.path.join(output_dir, f"ALM_Test_Results_{current_datetime}.xlsx")
matched_data.to_excel(output_file_path, index=False)

print(f"Test results saved to: {output_file_path}")
