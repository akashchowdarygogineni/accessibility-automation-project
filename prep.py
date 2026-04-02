import os
import atexit
import json
import random
import time
import socket
import subprocess
import shutil
import requests
import sys
import pandas as pd
from collections import OrderedDict
from openpyxl.styles import Alignment
from openpyxl import load_workbook
from requests.exceptions import Timeout

# =================================================
# CONFIGURATION
# =================================================


# JAR_PATH = r"C:\Users\akash\Downloads\pdfremediation-0.0.1-SNAPSHOT.jar"
#JAR_PATH = os.path.join(BASE_DIR, "pdfremediation-0.0.1-SNAPSHOT.jar")
# =================================================
# CONFIGURATION
# =================================================

version = sys.argv[1] if len(sys.argv) > 1 else "v1"
jar_name = sys.argv[2] if len(sys.argv) > 2 else None
DEFAULT_JAR_NAME = "pdfremediation-0.0.1-SNAPSHOT.jar"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

VERSION_FOLDER = os.path.join(BASE_DIR, version)


INPUT_FOLDER = os.path.join(VERSION_FOLDER, "pac_processed")

OUTPUT_FOLDER = os.path.join(VERSION_FOLDER, "prep_results")
SKIPPED_FOLDER = os.path.join(VERSION_FOLDER, "prep_skipped")
PROCESSED_FOLDER = os.path.join(VERSION_FOLDER, "prep_processed")

WORKING_FOLDER = os.path.join(VERSION_FOLDER, "working")

OUTPUT_EXCEL = os.path.join(OUTPUT_FOLDER, f"prep_final_summary_{version}.xlsx")

JAR_PATH = None

PORT = 8181
API_URL = f"http://localhost:{PORT}/pdfremediation/checker"
TOKEN = "8101d41ff43e329619f26ed53201ab9b7d0a123c"

for folder in [INPUT_FOLDER, OUTPUT_FOLDER, SKIPPED_FOLDER, PROCESSED_FOLDER]:
    os.makedirs(folder, exist_ok=True)

headers = {
    "Authorization": f"Token {TOKEN}",
    "Content-Type": "application/json"
}

jar_process = None

# =================================================
# SERVER CONTROL
# =================================================

def wait_for_server(port, timeout=60):
    start = time.time()
    while True:
        try:
            with socket.create_connection(("localhost", port), timeout=2):
                return True
        except OSError:
            if time.time() - start > timeout:
                return False
            time.sleep(2)

def start_jar():
    global jar_process
    print(":rocket: Starting JAR server...")
    jar_process = subprocess.Popen(
        ["java", "-jar", JAR_PATH, f"--server.port={PORT}"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    if not wait_for_server(PORT):
        print(":x: Server failed to start.")
        return False
    print(":white_check_mark: JAR server ready.\n")
    return True

def stop_jar():
    global jar_process
    if jar_process:
        print(":octagonal_sign: Stopping JAR server...")
        jar_process.terminate()
        jar_process.wait()
        time.sleep(2)

def delete_working_folder():
    if os.path.isdir(WORKING_FOLDER):
        try:
            shutil.rmtree(WORKING_FOLDER)
            print(":white_check_mark: Working folder deleted.")
        except Exception as e:
            print(f":warning: Could not delete working folder: {e}")

atexit.register(delete_working_folder)

# =================================================
# CHECK EXTRACTION
# =================================================

PRIORITY = {
    "Failed": 5,
    "Warning": 4,
    "Manual": 3,
    "Skipped": 2,
    "Passed": 1,
    "Info": 0
}

def map_status(status):
    if status == "Success":
        return "Passed"
    return status

def update_with_priority(checks, check_type, status):
    status = map_status(status)
    if status not in PRIORITY:
        return
    if check_type not in checks or PRIORITY[status] > PRIORITY[checks[check_type]]:
        checks[check_type] = status

def extract_checks(data, checks):
    if isinstance(data, dict):
        if (
            "type" in data and
            "status" in data and
            "checker_Standards" in data
        ):
            check_type = str(data["type"]).strip().lower()
            check_type = " ".join(check_type.split())
            status = data["status"]
            update_with_priority(checks, check_type, status)

        for key, value in data.items():
            if key.lower() == "error":
                continue
            if isinstance(value, (dict, list)):
                extract_checks(value, checks)

    elif isinstance(data, list):
        for item in data:
            extract_checks(item, checks)

# =================================================
# EXCEL
# =================================================

def rebuild_excel(all_rows):
    if not all_rows:
        print(":warning: No data to write.")
        return

    df = pd.DataFrame(all_rows)

    # Remove only this specific column
    column_to_remove = "1.2 time-based media"
    if column_to_remove in df.columns:
        df = df.drop(columns=[column_to_remove])

    df.to_excel(OUTPUT_EXCEL, index=False, sheet_name="Prep Accessibility Report")

    wb = load_workbook(OUTPUT_EXCEL)
    ws = wb.active

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 5
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(OUTPUT_EXCEL)
    print(":bar_chart: Excel Built Successfully\n")

# =================================================
# MAIN
# =================================================

if not os.path.isdir(INPUT_FOLDER):
    print(f":x: Input folder not found for version '{version}': {INPUT_FOLDER}")
    print("Create the folder and add PDFs, then run again.")
    raise SystemExit(0)

pdf_files = sorted([f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".pdf")])

if not pdf_files:
    print(f":warning: No PDF files found in: {INPUT_FOLDER}")
    print("Nothing to process. Add files in pac_processed and run again.")
    raise SystemExit(0)

# Resolve JAR path: explicit arg wins; otherwise auto-detect from version folder.
if jar_name:
    JAR_PATH = os.path.join(VERSION_FOLDER, jar_name)
else:
    jar_candidates = sorted(
        [f for f in os.listdir(VERSION_FOLDER) if f.lower().endswith(".jar")]
    )
    if DEFAULT_JAR_NAME in jar_candidates:
        JAR_PATH = os.path.join(VERSION_FOLDER, DEFAULT_JAR_NAME)
    elif len(jar_candidates) == 1:
        JAR_PATH = os.path.join(VERSION_FOLDER, jar_candidates[0])
    elif len(jar_candidates) > 1:
        print(f":x: Multiple JAR files found in: {VERSION_FOLDER}")
        print("Pass the jar name explicitly as the second argument.")
        print(f"Available JARs: {', '.join(jar_candidates)}")
        raise SystemExit(0)
    else:
        print(f":x: No JAR files found in: {VERSION_FOLDER}")
        print("Place a JAR in this version folder or pass jar name as second argument.")
        print(f"Example: python prep.py {version} {DEFAULT_JAR_NAME}")
        raise SystemExit(0)

if not os.path.isfile(JAR_PATH):
    print(f":x: JAR not found: {JAR_PATH}")
    print("Place the JAR in the selected version folder and pass jar name as second argument.")
    print(f"Example: python prep.py {version} {DEFAULT_JAR_NAME}")
    raise SystemExit(0)

print(f":file_folder: Using JAR: {JAR_PATH}")

if not start_jar():
    exit()

print(f":page_facing_up: Total PDFs Found: {len(pdf_files)}\n")

all_rows = []

for pdf_file in pdf_files:

    print(f":page_facing_up: Processing: {pdf_file}")

    input_path = os.path.join(INPUT_FOLDER, pdf_file)
    working_path = os.path.join(WORKING_FOLDER, pdf_file)
    processed_path = os.path.join(PROCESSED_FOLDER, pdf_file)
    skipped_path = os.path.join(SKIPPED_FOLDER, pdf_file)

    source_id = random.randint(10000, 99999)

    os.makedirs(WORKING_FOLDER, exist_ok=True)
    shutil.copy2(input_path, working_path)

    payload = {
        "inputpdf": working_path.replace("\\", "/"),
        "password": "",
        "level": "null",
        "source_id": source_id
    }

    try:
        print(":hourglass_flowing_sand: Waiting for JAR response (max 5 minutes)...")
        response = requests.post(API_URL, headers=headers, json=payload, timeout=300)
        response.raise_for_status()
        report = response.json()
        print(":white_check_mark: Response received!")

    except Timeout:
        print(":alarm_clock: Timeout! Copying file to skipped folder.")
        shutil.copy2(input_path, skipped_path)
        all_rows.append({"File Name": pdf_file, "Source ID": source_id})
        stop_jar()
        start_jar()
        continue

    except Exception as e:
        print(f":x: Error: {e}")
        shutil.copy2(input_path, skipped_path)
        all_rows.append({"File Name": pdf_file, "Source ID": source_id})
        stop_jar()
        start_jar()
        continue

    # Success
  
    for _ in range(5):
      try:
        shutil.copy2(input_path, processed_path)
        break
      except PermissionError:
        print("File still locked, retrying...")
        time.sleep(2)
    
    json_path = os.path.join(OUTPUT_FOLDER, os.path.splitext(pdf_file)[0] + f"_{version}.json")
  
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)

    checks = OrderedDict()
    extract_checks(report, checks)

    row = {"File Name": pdf_file, "Source ID": source_id, **checks}
    all_rows.append(row)

stop_jar()
rebuild_excel(all_rows)

# Delete working folder after the full run completes
print(":broom: Deleting working folder...")
if os.path.isdir(WORKING_FOLDER):
    try:
        shutil.rmtree(WORKING_FOLDER)
        print(":white_check_mark: Working folder deleted.")
    except Exception as e:
        print(f":warning: Could not delete working folder: {e}")
else:
    print(":information_source: Working folder already absent.")

print(":tada: ALL FILES COMPLETED")
print(":open_file_folder: Original files remain in input folder.")
print(":file_folder: Processed files copied to processed folder.")
print(":file_folder: Skipped files copied to skipped folder.")
print(f":bar_chart: Final Excel: {OUTPUT_EXCEL}")