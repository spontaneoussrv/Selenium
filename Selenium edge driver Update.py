import os
import sys
import time
import zipfile
import shutil
import ctypes
import stat
import subprocess
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# Utility: Logging Setup
# ============================================================
LOG_FILE = os.path.join(os.getcwd(), "selenium_install.txt")

def log(message):
    """Print to console and append to log file with timestamp."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ============================================================
# Step 0: Relaunch as Administrator (Safe Method)
# ============================================================
def run_as_admin():
    """Relaunch the script with Administrator privileges if not already elevated."""
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False

    if not is_admin:
        log("🔒 Requesting Administrator privileges...")
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{a}"' for a in sys.argv[1:]])
        args = f'"{script}" {params}'.strip()

        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, args, None, 1
        )
        sys.exit(0)


# ============================================================
# Step 1: Initialize
# ============================================================
if os.path.exists(LOG_FILE):
    os.remove(LOG_FILE)

log("==============================================")
log("     🧩 Microsoft Edge WebDriver Updater")
log("==============================================\n")
time.sleep(0.5)

run_as_admin()

# ============================================================
# Step 2: Detect SeleniumBasic Folder
# ============================================================
user_profile = os.environ.get("USERPROFILE", r"C:\Users\Default")
paths_to_check = [
    os.path.join(user_profile, "AppData", "Local", "SeleniumBasic"),
    os.path.join(user_profile, "AppData", "Local", "SleniumBasics"),
    r"C:\Program Files\SeleniumBasic"
]

selenium_basic_path = None

for path in paths_to_check:
    if os.path.exists(path):
        selenium_basic_path = path
        log(f"📁 Found SeleniumBasic folder: {path}")
        break

if selenium_basic_path is None:
    selenium_basic_path = paths_to_check[0]
    os.makedirs(selenium_basic_path, exist_ok=True)
    log(f"📁 No existing folder found — created new one: {selenium_basic_path}")

target_file = os.path.join(selenium_basic_path, "edgedriver.exe")


# ============================================================
# Step 3: Prepare Temp Folders
# ============================================================
download_dir = os.path.join(os.getcwd(), "edge_driver_downloads")
extract_dir = os.path.join(download_dir, "extracted")
os.makedirs(download_dir, exist_ok=True)

log("🗂️ Step 3: Temporary folders ready.")
log(f"   ➤ Download folder: {download_dir}")
log(f"   ➤ Extract folder : {extract_dir}\n")


# ============================================================
# Step 4: Launch Microsoft Edge
# ============================================================
log("🌐 Step 4: Launching Microsoft Edge browser...")
edge_options = Options()
edge_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True
})
driver = webdriver.Edge(options=edge_options)
driver.maximize_window()
log("   ✅ Edge launched successfully!")


# ============================================================
# Step 5: Open WebDriver Download Page
# ============================================================
log("🔗 Step 5: Opening official Edge WebDriver page...")
driver.get("https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH#downloads")
log("   🌍 Page loaded.")


# ============================================================
# Step 6: Click the x64 Download Button
# ============================================================
log("🖱️ Step 6: Finding and clicking the 'x64' download button...")
wait = WebDriverWait(driver, 30)
x64_button = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//span[contains(@class,'common-button-v1__content') and contains(.,'x64')]")
))
x64_button.click()
log("   ✅ Clicked the x64 download button.")


# ============================================================
# Step 7: Wait for File Download
# ============================================================
log("⬇️ Step 7: Waiting for the WebDriver ZIP to download...")
downloaded_file = None
timeout = time.time() + 120
spinner = ["|", "/", "-", "\\"]
i = 0

while time.time() < timeout:
    for file in os.listdir(download_dir):
        if file.endswith(".zip"):
            downloaded_file = os.path.join(download_dir, file)
            break
    if downloaded_file:
        break
    print(f"\r   ⏳ Waiting for file... {spinner[i % 4]}", end="", flush=True)
    i += 1
    time.sleep(0.5)

if not downloaded_file or not os.path.exists(downloaded_file):
    driver.quit()
    raise FileNotFoundError("\n❌ Download failed or no ZIP file found.")

log(f"   ✅ File downloaded: {downloaded_file}")


# ============================================================
# Step 8: Extract the ZIP File
# ============================================================
log("📦 Step 8: Extracting downloaded ZIP...")
os.makedirs(extract_dir, exist_ok=True)
with zipfile.ZipFile(downloaded_file, 'r') as zip_ref:
    zip_ref.extractall(extract_dir)
log(f"   ✅ Extracted to: {extract_dir}")


# ============================================================
# Step 9: Find .exe File Recursively
# ============================================================
log("🔍 Step 9: Searching for EdgeDriver executable...")
driver_exe = None
for root, _, files in os.walk(extract_dir):
    for file in files:
        if file.lower().endswith(".exe"):
            driver_exe = os.path.join(root, file)
            break
    if driver_exe:
        break

if not driver_exe or not os.path.exists(driver_exe):
    driver.quit()
    raise FileNotFoundError("❌ Could not find EdgeDriver executable inside ZIP.")

log(f"   ✅ Found EdgeDriver: {driver_exe}")


# ============================================================
# Step 10: Force Delete old edgedriver.exe
# ============================================================
log("🗑️ Step 10: Deleting old edgedriver.exe (if exists)...")

def force_delete(filepath):
    if os.path.exists(filepath):
        try:
            os.chmod(filepath, stat.S_IWRITE)
            os.remove(filepath)
            log("   ✅ Old edgedriver.exe deleted.")
        except PermissionError:
            log("   ⚠️ File in use — forcing termination...")
            subprocess.run(["taskkill", "/f", "/im", "edgedriver.exe"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(2)
            os.chmod(filepath, stat.S_IWRITE)
            os.remove(filepath)
            log("   ✅ Forced deletion completed.")
        except Exception as e:
            log(f"   ⚠️ Delete failed: {e}")
    else:
        log("   ℹ️ No old driver found.")

force_delete(target_file)


# ============================================================
# Step 11: Copy and Rename New Driver
# ============================================================
log("📤 Step 11: Copying new driver to SeleniumBasic folder...")
shutil.copy2(driver_exe, target_file)
log(f"   ✅ Installed new driver as: {target_file}")


# ============================================================
# Step 12: Cleanup Temporary Files
# ============================================================
log("🧹 Step 12: Cleaning up temporary files...")
try:
    if os.path.exists(downloaded_file):
        os.remove(downloaded_file)
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir)
    log("   ✅ Cleanup complete.")
except Exception as e:
    log(f"   ⚠️ Cleanup issue: {e}")


# ============================================================
# Step 13: Done
# ============================================================
log("🎉 Step 13: Process complete!")
log("   ➤ Edge WebDriver updated successfully.")
log(f"   ➤ Installed in: {selenium_basic_path}")
log("==============================================")
log("✅ Process Completed Successfully!")
log("==============================================\n")

time.sleep(2)
driver.quit()
