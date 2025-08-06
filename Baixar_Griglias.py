import os
import time
import json
import sys
import urllib.parse
import threading
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# === Global GUI Elements (to be updated by processing thread) ===
status_label = None
progress_bar = None
root_window = None # Reference to the main Tkinter window for updates

# === Configuration ===
DOWNLOAD_FOLDER = "Griglia"

# === Helper function to update GUI from any thread ===
def update_gui_status(message, style_name="TLabel"): # Changed 'color' to 'style_name'
    if status_label and root_window:
        # Schedule the update on the main Tkinter thread
        root_window.after(0, lambda: status_label.config(text=message, style=style_name))
        # Ensure the GUI updates immediately
        root_window.after(0, root_window.update_idletasks)

def update_gui_progress(value, maximum=None):
    if progress_bar and root_window:
        root_window.after(0, lambda: progress_bar.config(value=value, maximum=maximum if maximum is not None else progress_bar['maximum']))
        root_window.after(0, root_window.update_idletasks)




# === Load URLs from config.json (modified from your request) ===
def import_links_from_json():
    update_gui_status("Importing links from config.json...")
    try:
        with open('Griglia_links.json', 'r') as f:
            config_data = json.load(f)
        sharepoint_urls = []
        if 'SharePoint' in config_data:
            for key, value in config_data['SharePoint'].items():
                sharepoint_urls.append(value)
            update_gui_status(f"✅ Successfully imported {len(sharepoint_urls)} link(s).", "Success.TLabel")
            return sharepoint_urls
        else:
            update_gui_status("❌ ERROR: 'SharePoint' section not found in Griglia_lin.json.", "Error.TLabel")
            return []
    except FileNotFoundError:
        update_gui_status("❌ ERROR: 'Griglia_links.json' not found. Please create the configuration file.", "Error.TLabel")
        return []
    except json.JSONDecodeError as e:
        update_gui_status(f"❌ ERROR: Could not decode 'Griglia_links.json'. Check for JSON formatting errors: {e}", "Error.TLabel")
        return []
    except Exception as e:
        update_gui_status(f"❌ ERROR: An unexpected error occurred while loading Griglia_links: {e}", "Error.TLabel")
        return []
    



    
    




# === Extract model name from URL ===
def extract_model_name(full_url_entry):
    try:
        if '-' in full_url_entry:
            model_name = full_url_entry.split('-')[0].strip()
            update_gui_status(f"🏷️ Extracted model name: {model_name}")
            return model_name
        else:
            update_gui_status("⚠️ No model name found in URL entry format", "Warning.TLabel")
            return "Unknown"
    except Exception as e:
        update_gui_status(f"❌ Error extracting model name: {e}", "Error.TLabel")
        return "Unknown"

# === Get actual URL from entry ===
def extract_url_from_entry(full_url_entry):
    try:
        if '-' in full_url_entry:
            first_dash_index = full_url_entry.find('-')
            actual_url = full_url_entry[first_dash_index + 1:].strip()
            return actual_url
        else:
            return full_url_entry.strip()
    except Exception as e:
        update_gui_status(f"❌ Error extracting URL: {e}", "Error.TLabel")
        return full_url_entry

# === Wait for download and rename file ===
def wait_for_download_and_rename(download_path, model_name, original_filename, timeout=300):
    update_gui_status("⏳ Waiting for download to complete...")
    seconds = 0
    while seconds < timeout:
        update_gui_status("⏳ Processing {model_name} - {original_filename}")
        has_temp_files = any(f.endswith(('.crdownload', '.tmp')) for f in os.listdir(download_path))
        if not has_temp_files:
            time.sleep(2) # Give a moment for the file system to settle
            if not any(f.endswith(('.crdownload', '.tmp')) for f in os.listdir(download_path)): # Check again
                update_gui_status("✅ Download completed.", "Success.TLabel")

                try:
                    files_in_dir = os.listdir(download_path)
                    downloaded_file = None

                    if original_filename in files_in_dir:
                        downloaded_file = original_filename
                    else:
                        base_name = os.path.splitext(original_filename)[0]
                        extension = os.path.splitext(original_filename)[1]
                        for file in files_in_dir:
                            if base_name.lower() in file.lower() and file.endswith(extension):
                                downloaded_file = file
                                break
                        if not downloaded_file and files_in_dir:
                            files_with_time = [(f, os.path.getctime(os.path.join(download_path, f)))
                                               for f in files_in_dir
                                               if not f.startswith('.') and os.path.isfile(os.path.join(download_path, f))]
                            if files_with_time:
                                downloaded_file = max(files_with_time, key=lambda x: x[1])[0]

                    if downloaded_file:
                        old_path = os.path.join(download_path, downloaded_file)
                        file_extension = os.path.splitext(downloaded_file)[1]
                        file_name_without_ext = os.path.splitext(downloaded_file)[0]
                        new_filename = f"{model_name} - {file_name_without_ext}{file_extension}"
                        new_path = os.path.join(download_path, new_filename)

                        counter = 1
                        original_new_path_base, original_new_path_ext = os.path.splitext(new_filename)
                        while os.path.exists(new_path):
                            new_filename = f"{original_new_path_base} ({counter}){original_new_path_ext}"
                            new_path = os.path.join(download_path, new_filename)
                            counter += 1

                        if os.path.exists(old_path):
                            os.rename(old_path, new_path)
                            update_gui_status(f"📝 File renamed to: {os.path.basename(new_path)}")
                        else:
                            update_gui_status(f"⚠️ Could not find downloaded file: {downloaded_file}", "Warning.TLabel")
                    else:
                        update_gui_status("⚠️ Could not identify downloaded file for renaming", "Warning.TLabel")

                except Exception as rename_error:
                    update_gui_status(f"❌ Error renaming file: {rename_error}", "Error.TLabel")
                return True
        time.sleep(1)
        seconds += 1
    update_gui_status("⏱️ Download timed out.", "Warning.TLabel")
    return False

# === Create download directory ===
def create_download_directory():
    try:
        os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
        return os.path.abspath(DOWNLOAD_FOLDER)
    except Exception as e:
        update_gui_status(f"❌ Failed to create download directory: {e}", "Error.TLabel")
        return None

# === Setup Selenium WebDriver ===
def setup_webdriver(download_path):
    driver_path = os.path.join(os.getcwd(), "Driver", "msedgedriver.exe")
    if not os.path.exists(driver_path):
        update_gui_status(f"❌ ERROR: WebDriver not found at: {driver_path}", "Error.TLabel")
        return None

    edge_options = Options()
    edge_options.add_argument("--headless")
    edge_options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    edge_options.add_experimental_option("prefs", prefs)

    try:
        service = Service(executable_path=driver_path)
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.maximize_window()
        update_gui_status("✅ WebDriver initialized and maximized.", "Success.TLabel")
        return driver
    except Exception as e:
        update_gui_status(f"❌ WebDriver error: {e}", "Error.TLabel")
        return None

# === Handle Download Button Click ===
def handle_download_click(driver):
    # update_gui_status("📥 Attempting to click the download button...")
    download_button_selectors = [
        "//button[@data-automationid='downloadCommand']",
        "//button[contains(@aria-label, 'Download')]",
        "//button[contains(text(), 'Download')]",
        "//span[contains(text(), 'Download')]/parent::button"
    ]

    for selector in download_button_selectors:
        try:
            download_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, selector))
            )
            download_button.click()
            update_gui_status("⬇️ Download started (direct button).")
            return True
        except TimeoutException:
            continue

    update_gui_status("⚠️ Direct download button not found, trying 'More commands' button...", "Warning.TLabel")
    more_commands_selectors = [
        "//button[@title='More commands']",
        "//button[contains(@aria-label, 'More commands')]",
        "//button[contains(@aria-label, 'More actions')]",
        "//button[contains(@title, 'More')]",
        "//button[@data-automationid='moreCommands']",
        "//button[contains(text(), '...')]",
        "//i[contains(@class, 'ms-Icon--More')]/parent::button"
    ]

    for selector in more_commands_selectors:
        try:
            more_commands_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, selector))
            )
            more_commands_button.click()
            update_gui_status("✅ 'More commands' button clicked.")
            time.sleep(1) # Wait for dropdown to appear

            dropdown_download_selectors = [
                "//button[@data-automationid='downloadCommand']",
                "//div[contains(@class, 'ms-ContextualMenu')]//button[contains(@aria-label, 'Download')]",
                "//div[contains(@class, 'ms-ContextualMenu')]//button[contains(text(), 'Download')]",
                "//div[contains(@class, 'ms-ContextualMenu')]//span[contains(text(), 'Download')]/parent::button",
                "//ul[contains(@class, 'ms-ContextualMenu')]//button[contains(text(), 'Download')]",
                "//div[@role='menu']//button[contains(text(), 'Download')]"
            ]

            for download_selector in dropdown_download_selectors:
                try:
                    download_menu_item = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, download_selector))
                    )
                    download_menu_item.click()
                    update_gui_status("⬇️ Download started (via dropdown menu).")
                    return True
                except TimeoutException:
                    continue
            update_gui_status("❌ Could not find download option in dropdown menu.", "Error.TLabel")
            return False
        except TimeoutException:
            continue

    update_gui_status("❌ Could not find 'More commands' button or download option.", "Error.TLabel")
    return False

# === Download Files (Main processing loop) ===
def download_files_task(driver, url_entries, download_path):
    download_log = {}
    update_gui_status("\n📂 Starting file processing...")

    total_urls = len(url_entries)
    update_gui_progress(0, total_urls) # Initialize progress bar

    for i, url_entry in enumerate(url_entries):
        update_gui_status(f"\n🔄 Processing URL {i+1}/{total_urls}: {url_entry}")
        update_gui_progress(i + 1) # Update progress after each item starts processing
        try:
            model_name = extract_model_name(url_entry)
            actual_url = extract_url_from_entry(url_entry)

            url_base, file_name = actual_url.rsplit('/', 1)
            file_name = urllib.parse.unquote(file_name)

            update_gui_status(f"🏷️Processing  Model: {model_name}, 📁 File: {file_name}")

            # update_gui_status("🌐 Navigating to folder...")
            driver.get(url_base)
            time.sleep(5) # Give page some time to load, reduced from 10s for potentially faster feedback

            # update_gui_status("🔍 Looking for the file in the list...")
            list_items = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[role="row"]'))
            )

            matched = False
            for item in list_items:
                if file_name.strip() in item.text.strip():
                    # update_gui_status(f"✅ Found match: {item.text.strip()}")
                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", item)
                        time.sleep(1)
                        item.click()
                        matched = True
                        break
                    except Exception as click_error:
                        # update_gui_status(f"⚠️ Normal click failed, trying JS click: {click_error}", "Warning.TLabel")
                        try:
                            driver.execute_script("arguments[0].click();", item)
                            matched = True
                            break
                        except Exception as js_error:
                           print("Json erro") # update_gui_status(f"❌ JS click also failed: {js_error}", "Error.TLabel")

            if not matched:
                # update_gui_status(f"❌ File not found or not clickable: {file_name}", "Error.TLabel")
                download_log[url_entry] = "Not Found"
                continue

            time.sleep(2) # Wait after clicking file

            if handle_download_click(driver):
                if wait_for_download_and_rename(download_path, model_name, file_name):
                    download_log[url_entry] = "Success"
                else:
                    download_log[url_entry] = "Timeout"
            else:
                download_log[url_entry] = "Download Button Click Failed"

        except Exception as e:
            # update_gui_status(f"❌ Error processing file: {e}", "Error.TLabel")
            download_log[url_entry] = f"Error: {e}"

    return download_log

# === Save Log ===
def save_log_to_json(log_data):
    log_file = "download_log.json"
    update_gui_status(f"\n📝 Saving log to '{log_file}'...")
    try:
        with open(log_file, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, indent=4)
        update_gui_status("✅ Log saved.", "Success.TLabel")
    except Exception as e:
        update_gui_status(f"❌ Failed to save log: {e}", "Error.TLabel")

# === Main script logic to be run in a separate thread ===
def main_processing_thread():
    global root_window # Ensure we can access and close the main window
    update_gui_status("🚀 Starting SharePoint Downloader Script")

    url_entries = import_links_from_json()
    if not url_entries:
        update_gui_status("No URLs to process. Exiting.", "Error.TLabel")
        if root_window:
            root_window.after(3000, root_window.destroy) # Close after 3 seconds
        return

    download_directory = create_download_directory()

    if download_directory:
        driver = setup_webdriver(download_directory)
        if driver:
            log = download_files_task(driver, url_entries, download_directory)
            save_log_to_json(log)
            update_gui_status("🛑 Closing browser...", "Normal.TLabel")
            driver.quit()
            update_gui_status("✅ Done. Window will close in 3 seconds.", "Success.TLabel")
            if root_window:
                root_window.after(3000, root_window.destroy) # Close after 3 seconds
        else:
            update_gui_status("❌ WebDriver setup failed. Exiting.", "Error.TLabel")
            if root_window:
                root_window.after(3000, root_window.destroy)
    else:
        # update_gui_status("❌ Failed to create download directory. Exiting.", "Error.TLabel")
        if root_window:
            root_window.after(3000, root_window.destroy)

# === GUI Setup ===
def start_gui():
    global status_label, progress_bar, root_window

    root = tk.Tk()
    root_window = root # Assign to global variable
    root.title("DONWLOAD ALL GRIGLIA FILES")

    # Apply a modern theme
    style = ttk.Style(root)
    style.theme_use('vista') # Or 'clam' or 'alt' for cross-platform modern look

    # Configure styling for elements
    style.configure("TFrame", background="#f0f0f0")
    style.configure("TLabel", background="#f0f0f0", foreground="black", font=("Segoe UI", 9))

    # Define custom styles for different status colors
    style.configure("Error.TLabel", foreground="red", background="#f0f0f0")
    style.configure("Warning.TLabel", foreground="orange", background="#f0f0f0")
    style.configure("Success.TLabel", foreground="green", background="#f0f0f0")


    # Calculate window position to center it
    window_width = 450
    window_height = 120
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    root.resizable(False, False) # Prevent resizing
    root.attributes("-topmost", True) # Keep window on top


    frame = ttk.Frame(root, padding="15")
    frame.pack(expand=True, fill="both")

    status_label = ttk.Label(frame, text="Initializing...", wraplength=window_width - 30, style="TLabel") # Set initial style
    status_label.pack(pady=(0, 10), anchor="w")

    progress_bar = ttk.Progressbar(frame, orient="horizontal", length=window_width - 30, mode="determinate")
    progress_bar.pack(pady=(0, 5), anchor="w")

    # Start the main processing in a separate thread
    threading.Thread(target=main_processing_thread, daemon=True).start()

    root.mainloop() # Start the Tkinter event loop

# === Entry point for the script ===
if __name__ == "__main__":
    start_gui()