import subprocess
import os
import time
import threading
import psutil
import win32gui
import pywinauto
from pywinauto.application import Application
import re
import sys
import ctypes

WROBOT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "WRobot.exe")

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Restart the script with administrator privileges."""
    if is_admin():
        return True
    else:
        print("Requesting administrator privileges...")
        try:
            # Re-run the program with admin rights
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas", 
                sys.executable, 
                " ".join(sys.argv), 
                None, 
                1
            )
            return False
        except Exception as e:
            print(f"Failed to restart with admin privileges: {e}")
            return False

def find_process_by_name(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == process_name:
            return proc
    return None

def main():
    # Check for admin privileges at startup
    if not is_admin():
        print("This script requires administrator privileges to function properly.")
        if run_as_admin():
            return  # Exit current instance, admin version will start
        else:
            print("Failed to obtain administrator privileges. Exiting...")
            input("Press Enter to exit...")
            return
    
    print("Running with administrator privileges.")
    
    process_name = "WRobot.exe"
    while True:
        wrobot_process = find_process_by_name(process_name)

        if not wrobot_process:
            print(f"Process '{process_name}' not found. Starting it from: {WROBOT_PATH}")
            try:
                # Start WRobot with elevated privileges
                subprocess.Popen([WROBOT_PATH], shell=True)
                time.sleep(5)
                wrobot_process = find_process_by_name(process_name)
                if not wrobot_process:
                    print(f"Failed to start '{process_name}'. Retrying in 10 seconds...")
                    time.sleep(10)
                    continue
            except FileNotFoundError:
                print(f"Error: '{WROBOT_PATH}' not found. Please check the path.")
                time.sleep(10)
                continue
            except Exception as e:
                print(f"An error occurred while starting '{process_name}': {e}")
                time.sleep(10)
                continue

            print(f"Found {process_name} with PID: {wrobot_process.pid}")

            try:
                app = Application(backend="uia").connect(process=wrobot_process.pid)
                main_window = app.top_window()
                print("Successfully connected to the main window.")

                def click_button_by_id(automation_id):
                    try:
                        button = main_window.child_window(auto_id=automation_id, control_type="Button")
                        button.wait('visible', timeout=10)
                        button.click()
                        print(f"Button with auto_id '{automation_id}' clicked successfully.")
                    except Exception as e:
                        print(f"Could not find or click button with auto_id '{automation_id}'. Error: {e}")

                print("Waiting for server connection to be established...")
                start_time = time.time()
                timeout = 120

                def check_launch_bot_ready():
                    try:
                        button = main_window.child_window(auto_id="buttonLaunchBot", control_type="Button")
                        return button.exists() and button.is_enabled()
                    except Exception:
                        return False

                while time.time() - start_time < timeout:
                    if check_launch_bot_ready():
                        print("Server connection established! Button is now enabled.")
                        click_button_by_id("buttonLaunchBot")
                        break
                    time.sleep(1)
                else:
                    print("Timeout waiting for server connection to be established.")
                    continue

                print("Waiting for 'License Keys Management' window...")
                license_window = app.window(title="License Keys Management")
                license_window.wait('visible', timeout=30)
                print("Found 'License Keys Management' window.")
                login_button = license_window.child_window(auto_id="buttonLogin", control_type="Button")
                login_button.wait('visible', timeout=10)
                threading.Thread(target=login_button.click).start()
                print("Click on 'buttonLogin' (LOGIN) initiated in a separate thread.")

                print("Waiting for captcha window to appear...")
                time.sleep(1)
                try:
                    windows_found = pywinauto.findwindows.find_windows(title_re=".*MessageBox.*")
                    for hwnd in windows_found:
                        try:
                            win32gui.PostMessage(hwnd, 0x0010, 0, 0)  # WM_CLOSE
                        except: pass
                except: pass
                print("Captcha window handling complete.")

                print("Waiting for final WRobot window to appear...")
                time.sleep(3)

                final_window = None
                for window in app.windows():
                    try:
                        if "License" not in window.window_text() and window.is_visible():
                            final_window = window
                            break
                    except: continue

                if not final_window:
                    final_window = app.top_window()

                if final_window:
                    print(f"Found final window: '{final_window.window_text()}'")
                    final_window.set_focus()

                    if final_window.window_text() == "MessageBoxTrialVersion":
                        print("Detected trial message window with math equation.")
                        try:
                            trial_window = app.window(title="MessageBoxTrialVersion")
                            static_controls = trial_window.children(control_type="Text")
                            math_text = next((ctrl.window_text() for ctrl in static_controls if "+" in ctrl.window_text()), None)

                            if math_text:
                                match = re.search(r"(\d+)\s*\+\s*(\d+)", math_text)
                                if match:
                                    result = str(int(match.group(1)) + int(match.group(2)))

                                    input_field = trial_window.child_window(control_type="Edit")
                                    input_field.wait('visible', timeout=5)
                                    input_field.set_edit_text(result)
                                    print("Entered result into input field.")
                                    time.sleep(0.5)

                                    try:
                                        ok_button = trial_window.child_window(title="OK", control_type="Button")
                                        ok_button.wait('enabled', timeout=5)
                                        ok_button.click_input()
                                        print("Clicked OK button.")
                                    except Exception as e:
                                        print(f"Could not click OK: {e}")
                                else:
                                    print("Could not parse numbers from equation.")
                        except Exception as e:
                            print(f"An error occurred while solving the math captcha: {e}")
                    else:
                        print("Sending key combination 'Alt+C'...")
                        final_window.type_keys("%c")  # Alt+C
                        print("Key combination 'Alt+C' sent successfully.")
                else:
                    print("Final window not found.")

            except Exception as e:
                print(f"An error occurred during automation: {e}")
        else:
            print(f"'{process_name}' is running. Checking again in 5 seconds.")

        time.sleep(5)

if __name__ == "__main__":
    main()
