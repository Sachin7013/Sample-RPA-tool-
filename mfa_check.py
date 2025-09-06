import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


# Start Chrome
driver = webdriver.Chrome()

# Create evidence folder if not exists
evidence_dir = "evidence"
os.makedirs(evidence_dir, exist_ok=True)


driver.get("https://login.microsoftonline.com/")


WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "i0116"))
).send_keys("sachinbfrnd@gmail.com")

driver.find_element(By.ID, "idSIButton9").click()
time.sleep(5)


# Wait until page changes
WebDriverWait(driver, 10).until(lambda d: d.current_url != "https://login.microsoftonline.com/")

# Take screenshot for evidence
screenshot_path = os.path.join(evidence_dir, "login_page.png")
driver.save_screenshot(screenshot_path)

# Check MFA prompt
if ("Get a code to sign in" in driver.page_source or
    "Approve sign in request" in driver.page_source or
    "Enter code" in driver.page_source):
    print("✅ MFA is enabled — manual approval required")
    mfa_status = "MFA_ENABLED"
else:
    print("❌ MFA not detected — can continue automation")
    mfa_status = "MFA_NOT_ENABLED"

# Save evidence with status
final_screenshot = os.path.join(evidence_dir, f"{mfa_status}.png")
driver.save_screenshot(final_screenshot)
print(f"📸 Evidence saved: {final_screenshot}")


driver.quit()
