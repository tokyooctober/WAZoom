# start Zoom Window
# start recording

# Standard library imports
import re
import time
from datetime import datetime, timedelta

# Third-party library imports
import obsws_python as obs
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# ... existing code ...


def click_zoom_link(driver, zoom_link):
    try:
        # Find the element containing the Zoom link
        link_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, f'//a[contains(@href, "{zoom_link}")]')
            )
        )

        # Click the link
        link_element.click()
        print(f"Clicked on Zoom link: {zoom_link}")

        # Wait for a few seconds to allow the browser to process the click
        time.sleep(5)

    except Exception as e:
        print(f"An error occurred while clicking the Zoom link: {e}")


def minimize_native_zoom_window():
    try:
        # Wait for the Zoom main window to appear
        zoom_main_window = None
        attempts = 0
        max_attempts = 10

        while not zoom_main_window and attempts < max_attempts:
            zoom_windows = gw.getWindowsWithTitle("Zoom")
            if zoom_windows:
                zoom_main_window = zoom_windows[0]
            else:
                time.sleep(1)
                attempts += 1

        if not zoom_main_window:
            print("Zoom main window not found after waiting.")
            return

        # Find the dialog sub-window with no title
        dialog_window = None
        for window in gw.getAllWindows():
            if window.parent == zoom_main_window and not window.title:
                dialog_window = window
                break

        if not dialog_window:
            print("Dialog sub-window with no title not found.")
            return

        # Move the dialog window to a specific position (e.g., top-left corner)
        dialog_window.moveTo(0, 0)

        # Minimize the dialog window
        dialog_window.minimize()

        print("Dialog sub-window with no title has been moved and minimized.")
    except Exception as e:
        print(f"An error occurred while minimizing window: {e}")


# ... existing code ...


def wait_for_text_and_start_recording(driver, contact_name, target_text):
    try:
        # Open WhatsApp Web and navigate to the contact's chat
        driver.get("https://web.whatsapp.com/")
        print("WhatsApp Web opened. Please scan the QR code if needed.")

        # Wait for the chat list to load
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Chat list"]'))
        )

        # Search for the contact
        search_box = driver.find_element(
            By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'
        )
        search_box.send_keys(contact_name)

        # Wait for the contact to appear and click on it
        contact = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f'//span[@title="{contact_name}"]'))
        )
        contact.click()

        # Wait for the Zoom link to appear in the chat
        print(f"Waiting for zoom link in {contact_name}'s chat...")

        zoom_link_regex = (
            r"https://([A-Za-z0-9]+)\.zoom\.([A-Za-z0-9]+(/[A-Za-z0-9?=\.]+)+)"
        )
        zoom_link = None

        while not zoom_link:
            messages = driver.find_elements(
                By.XPATH, '//div[contains(@class, "message-in")]'
            )
            rmsg = list(reversed(messages))
            if rmsg:
                message = rmsg[0]
                print(message.text)
                match = re.search(zoom_link_regex, message.text)
                if match:
                    zoom_link = match.group(0)
                    break
            time.sleep(5)  # Wait 5 seconds before checking again

        print(f"Zoom link found: {zoom_link}")
        click_zoom_link(driver, zoom_link)

        print(f"Waiting for text: '{target_text}' in {contact_name}'s chat...")

        start_time = datetime.now()
        timeout_delta = timedelta(minutes=3)

        while True:
            if datetime.now() - start_time > timeout_delta:
                print(f"Timeout reached after 3 minutes")
                start_obs_recording()
                return

            messages = driver.find_elements(
                By.XPATH, '//div[contains(@class, "message-in")]'
            )
            reversed_messages = list(reversed(messages))
            if reversed_messages:
                message = reversed_messages[0]
                if target_text.lower() in message.text.lower():
                    print(f"Target text found: '{target_text}'")
                    start_obs_recording()
                    return  # Exit the loop after finding the target text
            time.sleep(5)  # Wait 5 seconds before checking again

        start_obs_recording()

    except Exception as e:
        print(f"An error occurred while waiting for text: {e}")
        return False


def start_obs_recording():
    try:
        client = obs.ReqClient()
        resp = client.get_version()
        # Access it's field as an attribute
        print(f"OBS Version: {resp.obs_version}")
        client.start_record()
        print("OBS recording started.")
        client.disconnect()

    except Exception as e:
        print(f"An error occurred while starting OBS recording: {e}")


# Call the function to open WhatsApp and get Zoom link
option = webdriver.ChromeOptions()
option.add_argument(r"user-data-dir=e:/src/WAZoom/whatsapp-cache")
option.add_argument("start-maximized")
service = Service(executable_path="e:\\src\\WAZoom\\chrome\\chromedriver.exe")
whatsapp_driver = webdriver.Chrome(service=service, options=option)
wait_for_text_and_start_recording(whatsapp_driver, "Edmund Trader Sharing", "JoinNow")
# wait_for_text_and_start_recording(whatsapp_driver, "mother HP", "JoinMow")

# Don't forget to close the drivers when you're done
whatsapp_driver.quit()

# ... existing code ...
