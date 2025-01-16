from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class LoginPage:
    def __init__(self, driver):
        self.driver = driver
        self.wait = WebDriverWait(driver, 10)

    # Locators
    login_baps_sso_button = (By.XPATH, "//button[contains(@label, 'Login with BAPS SSO')]")
    username_field = (By.XPATH, "//input[@name='userName']")
    password_field = (By.XPATH, "//input[@name='password']")
    sign_in_button = (By.XPATH, "//button[@class='btn btn-primary btn-submit']")
    dashboard_text = (By.XPATH, "//*[@id='main_content']/app-dashboard/div[1]/div/h4")

    # Actions
    def click_login_with_baps_sso(self):
        self.wait.until(EC.element_to_be_clickable(self.login_baps_sso_button)).click()

    def enter_username(self, username):
        self.wait.until(EC.visibility_of_element_located(self.username_field)).send_keys(username)

    def enter_password(self, password):
        self.wait.until(EC.visibility_of_element_located(self.password_field)).send_keys(password)

    def click_sign_in(self):
        self.wait.until(EC.element_to_be_clickable(self.sign_in_button)).click()

    def get_dashboard_text(self):
        return self.wait.until(EC.visibility_of_element_located(self.dashboard_text)).text.strip()
