import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

APP_URL = "https://streamlit.app"

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)

try:
    print(f"Visiting {APP_URL}...")
    driver.get(APP_URL)

    wait = WebDriverWait(driver, 20)

    # Wait until any button is present
    buttons = wait.until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "button"))
    )

    clicked = False

    for button in buttons:
        try:
            text = button.text.lower().strip()
            if "get this app back up" in text or "wake" in text:
                wait.until(EC.element_to_be_clickable(button)).click()
                print("Wake-up button clicked!")
                clicked = True
                break
        except Exception:
            continue

    if not clicked:
        print("App already awake or button not found.")

except Exception as e:
    print("Error:", str(e))

finally:
    driver.quit()
