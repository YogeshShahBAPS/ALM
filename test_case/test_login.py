import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utilities.excel import get_matched_test_cases
from datetime import datetime

test_data = get_matched_test_cases()
print(test_data)
# Initialize the WebDriver
driver = webdriver.Chrome()
driver.maximize_window()

# Create a WebDriverWait instance
wait = WebDriverWait(driver, 10)

# Loop through each row of the DataFrame to perform login tests
for index, row in test_data.iterrows():
    tc_number = row['MTC']
    user_name = row['username']  # Get username from the Excel file
    password = row['password']   # Get password from the Excel file
    expected_text = row['Expected Text']  # Expected dashboard text

    result = "FAIL"  # Default result

    try:
        # Navigate to the login page
        driver.get("https://qa.baps.dev/alm/signIn")

        # Wait for the 'Login with BAPS SSO' button and click it
        loginbaps_sso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@label, 'Login with BAPS SSO')]"))).click()

        # Wait for username and password fields and input credentials
        username_field = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='userName']"))).send_keys(user_name)
        password_field = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='password']"))).send_keys(password)

        # Click the sign-in button
        sign_in_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='btn btn-primary btn-submit']"))).click()

        # Check if the Dashboard text is present
        dashboard_xpath = "//*[@id='main_content']/app-dashboard/div[1]/div/h4"
        actual_text = wait.until(EC.visibility_of_element_located((By.XPATH, dashboard_xpath))).text

        # Compare actual text with the expected text
        if actual_text.strip() == expected_text.strip():
            result = "PASS"
        else:
            result = "FAIL"

    except Exception as e:
        print(f"Error during test case {tc_number}: {str(e)}")
        result = "FAIL"  # If there's any error during the process, mark it as FAIL

    # Add the result to the DataFrame in a new column 'Result'
    test_data.at[index, 'Result'] = result

# Close the driver
driver.quit()

# Generate the new filename with the current date and time
current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file_path = r"C:\Users\Tmp Admin\PycharmProjects\SeleniumPython\PythonProject\HTML Report\ALM_Test_Results.xlsx"

# Save the updated DataFrame to the new Excel file
test_data.to_excel(output_file_path, index=False)

print(f"Test results saved to a new Excel file: {output_file_path}")
