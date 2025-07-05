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
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui

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


def move_zoom_dialog_offscreen(driver):
    # Get all windows and analyze their properties
    all_windows = gw.getAllWindows()
    print("Searching for VideoFrameWnd window...")
    search_term = "videoframewnd"
    # Find window at specific coordinates
    video_port_windows = [w for w in all_windows if search_term in w.title.lower()]

    if not video_port_windows:
        print("No window with title 'VideoFrameWnd' found")

        return

    print(f"Found {len(video_port_windows)} VideoFrameWnd window(s)")

    # Get the video window
    video_port_window = video_port_windows[0]

    try:
        # Get current window position and size
        current_x = video_port_window.left
        current_y = video_port_window.top
        window_width = video_port_window.width
        window_height = video_port_window.height

        print(f"Current window position: ({current_x}, {current_y})")
        print(f"Window size: {window_width} x {window_height}")

        # Calculate target position (off-screen)
        target_x = 1413
        target_y = -1000

        drag_x = current_x + (window_width // 2)  # Center horizontally
        drag_y = current_y + (window_height // 2)  # Center of title bar

        print(f"Will drag from position: ({drag_x}, {drag_y})")
        print(f"Will drag to position: ({target_x}, {target_y})")

        # Ensure window is active/focused first
        video_port_window.activate()
        time.sleep(0.5)  # Wait for window to activate

        # Perform the drag operation
        # Move to the drag point
        pyautogui.moveTo(drag_x, drag_y, duration=0.5)
        time.sleep(0.2)

        # Press and hold left mouse button
        pyautogui.mouseDown()
        time.sleep(0.2)

        # Drag to target position
        pyautogui.dragTo(target_x, target_y, duration=1.0)
        time.sleep(0.2)

        # Release mouse button
        pyautogui.mouseUp()

        print(
            f"VideoFrameWnd window has been dragged to position ({target_x}, {target_y})"
        )

        # Verify the new position
        time.sleep(0.5)
        video_port_window = gw.getWindowsWithTitle(video_port_window.title)[0]
        new_x = video_port_window.left
        new_y = video_port_window.top
        print(f"New window position: ({new_x}, {new_y})")

    except Exception as e:
        print(f"Error moving zoom dialog off-screen: {e}")


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

        # Move the Zoom dialog window off-screen
        time.sleep(10)  # wait 10 seconds to allow the zoom window to load
        move_zoom_dialog_offscreen(driver)

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
