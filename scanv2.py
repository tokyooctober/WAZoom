# Standard library imports
import re
import time
import logging
from datetime import datetime, timedelta
import traceback

# Setup logging with date-time filename
log_filename = f"wazoom_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler()  # Also log to console
    ]
)
logger = logging.getLogger(__name__)

# Third-party library imports
import obsws_python as obs
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui
from ctypes import cast, POINTER
try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from comtypes import CLSCTX_ALL
    AUDIO_AVAILABLE = True
except ImportError:
    logger.warning("pycaw and comtypes libraries not available. Audio control functions will not work.")
    logger.warning("Install with: pip install pycaw comtypes")
    AUDIO_AVAILABLE = False


def set_windows_volume(volume_level):
    """
    Set Windows system volume to a specified level.
    
    Args:
        volume_level (int): Volume level from 0 to 100 (0 = mute, 100 = max volume)
    
    Returns:
        bool: True if successful, False otherwise
    """
    if not AUDIO_AVAILABLE:
        logger.error("Audio control libraries not available. Install pycaw and comtypes.")
        return False
        
    try:
        # Validate and clamp input parameter
        if not isinstance(volume_level, (int, float)):
            logger.error(f"Volume level must be a number. Got: {volume_level}")
            return False
        
        # Clamp volume level to valid range
        volume_level = max(0, min(100, int(volume_level)))
        
        # Get the default audio device
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        # Convert volume level from 0-100 scale to 0.0-1.0 scale
        volume_float = volume_level / 100.0
        
        # Set the volume using correct method name
        volume.SetMasterVolumeLevelScalar(volume_float, None)
        
        logger.info(f"Windows volume set to {volume_level}%")
        return True
        
    except Exception as e:
        logger.error(f"Error setting Windows volume: {e}")
        logger.error(traceback.format_exc())
        return False


def get_windows_volume():
    """
    Get the current Windows system volume level.
    
    Returns:
        int: Current volume level from 0 to 100, or -1 if error
    """
    if not AUDIO_AVAILABLE:
        logger.error("Audio control libraries not available. Install pycaw and comtypes.")
        return -1
        
    try:
        # Get the default audio device
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        # Get current volume (returns float between 0.0 and 1.0)
        current_volume = volume.GetMasterVolumeLevelScalar()
        
        # Convert to 0-100 scale and round to nearest integer
        volume_percentage = round(current_volume * 100)
        
        logger.info(f"Current Windows volume: {volume_percentage}%")
        return volume_percentage
        
    except Exception as e:
        logger.error(f"Error getting Windows volume: {e}")
        logger.error(traceback.format_exc())
        return -1


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
        logger.info(f"Clicked on Zoom link: {zoom_link}")

        # Wait for a few seconds to allow the browser to process the click
        time.sleep(5)

    except Exception as e:
        logger.error(f"An error occurred while clicking the Zoom link: {e}")


def move_zoom_dialog_offscreen():
    # Get all windows and analyze their properties
    all_windows = gw.getAllWindows()
    logger.info("Searching for VideoFrameWnd window...")
    search_term = "videoframewnd"
    # Find window at specific coordinates
    video_port_windows = [w for w in all_windows if search_term in w.title.lower()]

    if not video_port_windows:
        logger.warning("No window with title 'VideoFrameWnd' found")

        return

    logger.info(f"Found {len(video_port_windows)} VideoFrameWnd window(s)")

    # Get the video window
    video_port_window = video_port_windows[0]

    try:
        # Get current window position and size
        current_x = video_port_window.left
        current_y = video_port_window.top
        window_width = video_port_window.width
        window_height = video_port_window.height

        logger.info(f"Current window position: ({current_x}, {current_y})")
        logger.info(f"Window size: {window_width} x {window_height}")

        # Calculate target position (off-screen)
        target_x = 1413
        target_y = -1000

        drag_x = current_x + (window_width // 2)  # Center horizontally
        drag_y = current_y + (window_height // 2)  # Center of title bar

        logger.info(f"Will drag from position: ({drag_x}, {drag_y})")
        logger.info(f"Will drag to position: ({target_x}, {target_y})")

        # Ensure window is active/focused first
        logger.info("Activating window")
        video_port_window.activate()
        time.sleep(0.5)  # Wait for window to activate

        # Perform the drag operation
        # Move to the drag point
        logger.info("Moving to drag point")
        pyautogui.moveTo(drag_x, drag_y, duration=0.5)
        time.sleep(0.2)

        # # Press and hold left mouse button
        # pyautogui.mouseDown()
        # time.sleep(0.2)

        # Drag to target position
        logger.info("Dragging to target position")
        pyautogui.dragTo(target_x, target_y, duration=1.0, button="left")
        time.sleep(0.2)

        # # Release mouse button
        # pyautogui.mouseUp()

        logger.info(
            f"VideoFrameWnd window has been dragged to position ({target_x}, {target_y})"
        )

        # Verify the new position
        time.sleep(0.5)
        video_port_window = gw.getWindowsWithTitle(video_port_window.title)[0]
        new_x = video_port_window.left
        new_y = video_port_window.top
        logger.info(f"New window position: ({new_x}, {new_y})")

    except Exception as e:
        logger.error(f"Error moving zoom dialog off-screen: {e}")


def get_latest_available_incoming_date(driver):
    """
    Inspect recent incoming messages for any available date. If none carry a date,
    try detecting a day divider like 'Today' in the chat. Returns a date or None.
    """
    try:
        # Find all elements with data-pre-plain-text (these contain timestamps)
        messages = driver.find_elements(By.CSS_SELECTOR, "[data-pre-plain-text]")

        timestamps = []
        for msg in messages:
            raw = msg.get_attribute("data-pre-plain-text")
            # Format: [10/6/24, 10:14 AM] Luke Woo:
            match = re.match(r"\[(.*?)\]", raw)
            if match:
                try:
                    timestamp_str = match.group(1)
                    dt = datetime.strptime(timestamp_str, "%I:%M %p, %m/%d/%Y")
                    timestamps.append(dt)
                except ValueError:
                    logger.warning(f"Invalid timestamp format: {timestamp_str} ")
                    pass  # Skip if format is wrong

        # Get the latest datetime
        if timestamps:
            latest = max(timestamps)
            logger.info(f"Latest message timestamp: {latest.strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            logger.warning("No valid timestamps found.")
        return latest
    except Exception as e:
        logger.error(f"An error occurred while getting the latest available incoming date: {e}")
        logger.error(traceback.format_exc())
        return None



def wait_for_text_and_start_recording(driver, contact_name, target_text):
    try:
        # Open WhatsApp Web and navigate to the contact's chat
        driver.get("https://web.whatsapp.com/")
        logger.info("WhatsApp Web opened. Please scan the QR code if needed.")

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
        
        # Wait until the latest available incoming message date is today
        while True:
            last_message_date = get_latest_available_incoming_date(driver)
            logger.info(f"last message date: {last_message_date}")
            # logger.debug(f"now date is      : {datetime.now().date()}")
            # logger.debug(f"last message date day: {last_message_date.day}")
            # logger.debug(f"now date day is      : {datetime.now().day}")
            if last_message_date is not None and last_message_date.day == datetime.now().day:
                break
            
            time.sleep(5)

        # Wait for the Zoom link to appear in the chat
        logger.info(f"Waiting for zoom link in {contact_name}'s chat...")

        zoom_link_regex = (
            r"https://([A-Za-z0-9]+)\.zoom\.([A-Za-z0-9]+(/[A-Za-z0-9?=\.]+)+)"
        )
        zoom_link = None

        messages = driver.find_elements(
            By.XPATH, '//div[contains(@class, "message-in")]'
        )


        while not zoom_link:
            messages = driver.find_elements(
                By.XPATH, '//div[contains(@class, "message-in")]'
            )
            rmsg = list(reversed(messages))
            if rmsg:
                message = rmsg[0]
                # Log the message text and its date/time
                logger.info(f"Message: {message.text}")
                    
                match = re.search(zoom_link_regex, message.text)
                if match:
                    zoom_link = match.group(0)
                    break
            time.sleep(5)  # Wait 5 seconds before checking again

        logger.info(f"Zoom link found: {zoom_link}")
        click_zoom_link(driver, zoom_link)

        set_windows_volume(25)
        
        logger.info(f"Waiting for text: '{target_text}' in {contact_name}'s chat...")

        start_time = datetime.now()
        minutes=3
        timeout_delta = timedelta(minutes=minutes)

        while True:
            if datetime.now() - start_time > timeout_delta:
                logger.warning(f"Timeout reached after '{minutes}' minutes")
                # start_obs_recording()
                break  # quit this loop

            messages = driver.find_elements(
                By.XPATH, '//div[contains(@class, "message-in")]'
            )
            reversed_messages = list(reversed(messages))
            if reversed_messages:
                message = reversed_messages[0]

                if target_text.lower() in message.text.lower():
                    logger.info(f"Target text found: '{target_text}' from today")
                    # start_obs_recording() 
                    break
            time.sleep(5)  # Wait 5 seconds before checking again

        # Move the Zoom dialog window off-screen
        move_zoom_dialog_offscreen()

        start_obs_recording()

    except Exception as e:
        logger.error(f"An error occurred while waiting for text: {e}")
        logger.error(traceback.format_exc())
        return False


def start_obs_recording():
    try:
        client = obs.ReqClient()
        resp = client.get_version()
        # Access it's field as an attribute
        logger.info(f"OBS Version: {resp.obs_version}")
        client.start_record()
        logger.info("OBS recording started.")
        client.disconnect()

    except Exception as e:
        logger.error(f"An error occurred while starting OBS recording: {e}")


def main():
    """
    Main function to run the WhatsApp Zoom automation.
    """
    try:
        # Call the function to open WhatsApp and get Zoom link
        option = webdriver.ChromeOptions()
        option.add_argument(r"user-data-dir=e:/src/WAZoom/whatsapp-cache")
        option.add_argument("start-maximized")
        service = Service(executable_path="e:\\src\\WAZoom\\chrome\\chromedriver.exe")
        whatsapp_driver = webdriver.Chrome(service=service, options=option)
        
        logger.info("Starting WhatsApp Zoom automation...")
        wait_for_text_and_start_recording(whatsapp_driver, "Edmund Trader Sharing", "Join Now")
        # wait_for_text_and_start_recording(whatsapp_driver, "mother HP", "JoinMow")
        
        # Don't forget to close the drivers when you're done
        whatsapp_driver.quit()
        logger.info("WhatsApp Zoom automation completed successfully.")
        
    except Exception as e:
        logger.error(f"An error occurred in main function: {e}")
        logger.error(traceback.format_exc())


if __name__ == "__main__":
    main()
