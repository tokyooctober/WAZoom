from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import obsws_python as obs
import time
import re


def main():
    # Setup Selenium WebDriver
    option = webdriver.ChromeOptions()
    option.add_argument(r"user-data-dir=c:\temp\whatsapp")
    # option.add_experimental_option("excludeSwitches", ["enable-automation"])
    # option.add_experimental_option('useAutomationExtension', False)
    option.add_argument("start-maximized")
    #service = Service(ChromeDriverManager().install())
    service = Service(executable_path='e:\src\zoom\chrome\chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=option)

    try:
        # Open WhatsApp Web
        driver.get('https://web.whatsapp.com')
        print("Scan the QR code to log in to WhatsApp Web.")

        # Wait until the page is loaded and QR code is scanned
        WebDriverWait(driver, 600).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'span[dir="ltr"]'))
            
        )
        print("Logged into WhatsApp Web.")

        # wait for WhatsApp to load completely
        #time.sleep(5)
       # driver.find_element(By.CSS_SELECTOR, "span[title='Edmund Trader Sharing']").click()
       # driver.find_element(By.CSS_SELECTOR, "span[title='mother HP']").click()

        foundZoom=False
        while True:
            # Wait for a new message
            # message = WebDriverWait(driver, 600).until(
            #     EC.presence_of_element_located((By.CSS_SELECTOR, 'span._ao3e[dir="ltr"]'))
            # )
            elements = driver.find_elements(By.CSS_SELECTOR, 'span._ao3e[dir="ltr"]')
            if not foundZoom: 
                findZoomLink(elements,driver)
                foundZoom=True
            if foundZoom:
                if findJoinNow(elements):
                    print("Starting recording...")
                    start_recording()
                    break
            # Sleep for a short duration before checking for new messages again
            time.sleep(5)

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        print("Quiting...")
        driver.quit()


def findJoinNow(elements):
    print("=====")
    found=False
    for element in elements:
        print(f"Text:{element.text}")
        # Get the message text
        message_text = element.text
        if not isMatch(message_text, "JoinNow"):
            continue
        found=True
    return found

def findZoomLink(elements, driver):
    print("------")
    found=False
    for element in elements:
        print(f"Text:{element.text}")
        # Get the message text
        message_text = element.text
        if not isMatch(message_text, "edmund lee"):
            continue
        found=True
        # Check if the message contains a Zoom link
        zoom_link = extract_zoom_link(message_text)
        if zoom_link:
            print(f"Zoom link found:>{zoom_link}<")
            # Click on the Zoom link
            driver.execute_script("window.open('');")
            driver.switch_to.window(driver.window_handles[1])
            driver.get(zoom_link)
            driver.switch_to.window(driver.window_handles[0])
    return found

def isMatch(text, pattern):
    match = re.search(pattern,text, re.IGNORECASE)
    return True if match else False

def extract_zoom_link(text):
    # Regex pattern to match Zoom meeting links
    zoom_link_pattern = r"https://([A-Za-z0-9]+)\.zoom\.([A-Za-z0-9]+(/[A-Za-z0-9?=\.]+)+)"

    match = re.search(zoom_link_pattern, text)
    return match.group(0) if match else None

def start_recording():
    # client = obs.ReqClient(host="localhost", port=4455, password="strongpwd", timeout=3)
    client=obs.ReqClient()
    resp = client.get_version()
    # Access it's field as an attribute
    print(f"OBS Version: {resp.obs_version}")
    client.start_record()
    client.disconnect()

if __name__ == "__main__":
    main()
