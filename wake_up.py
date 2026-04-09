import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time

# Replace this with your actual Streamlit URL
APP_URL = "https://streamlit.app"

options = Options()
options.add_argument("--headless") # Runs without a visible window
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)

try:
    print(f"Visiting {APP_URL}...")
    driver.get(APP_URL)
    time.sleep(10) # Wait for the page to load
    
    # This looks for the "Yes, get this app back up!" button
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for button in buttons:
        if "get this app back up" in button.text.lower():
            button.click()
            print("Button found and clicked! Waking up...")
            time.sleep(5)
            break
    else:
        print("App was already awake!")
finally:
    driver.quit()
