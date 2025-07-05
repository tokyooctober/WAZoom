import pygetwindow as gw
import pyautogui
import time
import sys

# Configure pyautogui for safety
pyautogui.FAILSAFE = True  # Move mouse to corner to abort
pyautogui.PAUSE = 0.1  # Small pause between actions


def safe_print(text):
    """Safely print text that may contain Unicode characters"""
    try:
        print(text)
    except UnicodeEncodeError:
        # Fallback: encode as utf-8 and decode with error handling
        try:
            encoded = text.encode("utf-8", errors="replace").decode("utf-8")
            print(encoded)
        except:
            print(f"[Unicode text that couldn't be displayed: {repr(text)}]")


def move_zoom_dialog_offscreen():
    try:
        # Set console encoding to UTF-8 if possible
        if sys.platform == "win32":
            try:
                import codecs

                sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
            except:
                pass  # Fallback to safe printing

        # Get all windows and analyze their properties
        all_windows = gw.getAllWindows()

        for window in all_windows:
            safe_print(f"Title: {window.title}")
            safe_print(f"Size: {window.size}")
            safe_print(f"Position: {window.left}, {window.top}")
            safe_print(f"Visible: {window.visible}")
            safe_print("---")

        search_term = "videoframewnd"
        # Find window at specific coordinates
        video_port_windows = [w for w in all_windows if search_term in w.title.lower()]

        if not video_port_windows:
            safe_print("No window with title 'VideoFrameWnd' found")
            safe_print("Available windows with similar names:")
            for w in all_windows:
                if "video" in w.title.lower() or "frame" in w.title.lower():
                    safe_print(f"  - {w.title}")
            return

        safe_print(f"Found {len(video_port_windows)} VideoFrameWnd window(s)")

        # Debug: Show the titles of found windows
        for i, window in enumerate(video_port_windows):
            safe_print(f"  Window {i+1}: '{window.title}'")

        # Get the video window
        video_port_window = video_port_windows[0]

        # Try multiple approaches to move the window
        success = False
        max_attempts = 3

        for attempt in range(max_attempts):
            safe_print(f"\nAttempt {attempt + 1}/{max_attempts}")

            try:
                # Get current window position and size
                current_x = video_port_window.left
                current_y = video_port_window.top
                window_width = video_port_window.width
                window_height = video_port_window.height

                safe_print(f"Current window position: ({current_x}, {current_y})")
                safe_print(f"Window size: {window_width} x {window_height}")

                # Calculate target position (off-screen)
                target_x = 1413
                target_y = -1000

                drag_x = current_x + (window_width // 2)  # Center horizontally
                drag_y = current_y + (window_height // 2)  # Center of title bar

                safe_print(f"Will drag from position: ({drag_x}, {drag_y})")
                safe_print(f"Will drag to position: ({target_x}, {target_y})")

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

                safe_print(
                    f"VideoFrameWnd window has been dragged to position ({target_x}, {target_y})"
                )

                # Verify the new position
                time.sleep(0.5)
                video_port_window = gw.getWindowsWithTitle(video_port_window.title)[0]
                new_x = video_port_window.left
                new_y = video_port_window.top
                safe_print(f"New window position: ({new_x}, {new_y})")

                # Check if the move was successful
                if abs(new_x - target_x) < 50 and abs(new_y - target_y) < 50:
                    safe_print("Window move successful!")
                    success = True
                    break
                else:
                    safe_print(f"Window move may not have been successful. Retrying...")
                    time.sleep(1)

            except Exception as e:
                safe_print(f"Error in attempt {attempt + 1}: {e}")
                if attempt < max_attempts - 1:
                    time.sleep(1)
                    continue

        if not success:
            safe_print(
                "Failed to move window after all attempts. Trying alternative method..."
            )

    except Exception as e:
        safe_print(f"An error occurred while moving window off-screen: {e}")


if __name__ == "__main__":
    # Add a small delay to allow time for the dialog to appear
    time.sleep(2)

    # Call the function to move dialog off screen
    move_zoom_dialog_offscreen()
