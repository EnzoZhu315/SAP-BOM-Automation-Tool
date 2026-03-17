import os
import ctypes
import shutil
import subprocess
import time
import ssl
import win32com.client
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ==============================================================================
#  Part 1: Environment Initialization (DLLs, Network, Credentials)
# ==============================================================================
print(" --- Initializing Runtime Environment ---")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Note: Python_Engine should contain necessary DLLs for win32com in portable mode
dll_dir = os.path.join(BASE_DIR, "Python_Engine")

try:
    # Use generic loading for portable environment support
    ctypes.windll.kernel32.LoadLibraryW(os.path.join(dll_dir, "pywintypes313.dll"))
    ctypes.windll.kernel32.LoadLibraryW(os.path.join(dll_dir, "pythoncom313.dll"))
    print("System Components (DLL) Loaded")
except Exception as e:
    print(f"DLL Loading Info: {e}")

# Secure connection settings
os.environ['NO_PROXY'] = 'google.com,googleapis.com,sheets.googleapis.com'
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except:
    pass

# Localize Credentials (Handles access restrictions from network drives)
# [DE-IDENTIFIED] Replace the source path with your environment's path
SOURCE_JSON_PATH = r"YOUR_SHARED_DRIVE_PATH\service_account.json"
LOCAL_JSON_PATH = os.path.join(os.environ['TEMP'], "temp_credential_cache.json")

if os.path.exists(SOURCE_JSON_PATH):
    shutil.copy2(SOURCE_JSON_PATH, LOCAL_JSON_PATH)
    CREDENTIALS_JSON = LOCAL_JSON_PATH
else:
    CREDENTIALS_JSON = os.path.join(BASE_DIR, "service_account.json")

# ==============================================================================
# Part 2: Configuration (Masked Information)
# ==============================================================================
# [DE-IDENTIFIED] Placeholder for SAP credentials and system IDs
USER_INFO = {"user": "YOUR_USERNAME", "password": "YOUR_PASSWORD"}
SAP_SYSTEM_ID = "YOUR_SAP_SYSTEM_CONNECTION_NAME"
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
GS_KEY = 'YOUR_GOOGLE_SHEET_UNIQUE_ID_HERE'

# ==============================================================================
# Part 3: Google Sheets Logic
# ==============================================================================
def get_p_tasks_from_gs():
    print(f"\n --- Checking Google Sheets for Tasks ---")
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_JSON, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(GS_KEY)

        # Access target worksheet
        sheet = spreadsheet.worksheet("Task_List_Sheet")

        all_data = sheet.get_all_values()
        tasks = []

        for i, row in enumerate(all_data):
            if i == 0: continue  # Skip Header

            p_no = str(row[0]).strip()
            status = row[1] if len(row) > 1 else ""

            # Filter logic: Starts with P and status is not success
            if p_no.upper().startswith('P') and status.lower() != "success":
                tasks.append({"p_no": p_no, "row_idx": i + 1})

        print(f"Tasks Found: {len(tasks)}")
        return tasks, sheet
    except Exception as e:
        print(f" Google Sheet Access Failed: {e}")
        return [], None

# ==============================================================================
# Part 4: SAP Transaction Logic (CS01/CS02)
# ==============================================================================
def get_sap_session():
    print("\n🚀 --- Connecting to SAP ---")
    try:
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except:
            subprocess.Popen(SAP_LOGON_PATH)
            time.sleep(10)
            SapGuiAuto = win32com.client.GetObject("SAPGUI")

        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(SAP_SYSTEM_ID, True)
        session = connection.Children(0)

        # Login handling
        if session.findById("wnd[0]/usr/txtRSYST-BNAME", False):
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USER_INFO['user']
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = USER_INFO['password']
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
            # Handle multi-logon warning if present
            try:
                if "wnd[1]" in str(session.ActiveWindow.Name):
                    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                pass
        print("SAP Logon Successful")
        return session
    except Exception as e:
        print(f" SAP Connection Failed: {e}")
        return None

def run_sap_bom_maintenance(session, p_number):
    print(f"   Processing Material: {p_number}")
    try:
        # 1. Enter Transaction CS01
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS01"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # 2. Input Header Information
        session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = p_number
        session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "YOUR_PLANT_CODE"
        session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
        session.findById("wnd[0]").sendVKey(0)

        # Bypass initial warnings
        for _ in range(3):
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)

        # 3. Component Maintenance Loop
        # [DE-IDENTIFIED] Masked specific component lists
        table = "wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT"
        items = [
            ("L", "COMP_ID_A", "1.0", "PC"),
            ("L", "COMP_ID_B", "2.0", "G"),
        ]

        for i, (ict, comp, qty, unit) in enumerate(items):
            session.findById(f"{table}/ctxtRC29P-POSTP[1,{i}]").text = ict
            session.findById(f"{table}/ctxtRC29P-IDNRK[2,{i}]").text = comp
            session.findById(f"{table}/txtRC29P-MENGE[4,{i}]").text = qty
            session.findById(f"{table}/ctxtRC29P-MEINS[5,{i}]").text = unit
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)

        # 4. Handle Status Warnings & Save
        for _ in range(3):
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)

        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(1.5)
        print(f"       BOM Saved Successfully")
        return True

    except Exception as e:
        print(f"       Transaction Failed: {e}")
        return False

# ==============================================================================
#  Part 5: Main Entry Point
# ==============================================================================
if __name__ == "__main__":
    tasks, sheet_obj = get_p_tasks_from_gs()

    if tasks and sheet_obj:
        sap_session = get_sap_session()
        if sap_session:
            for item in tasks:
                if run_sap_bom_maintenance(sap_session, item['p_no']):
                    try:
                        sheet_obj.update_cell(item['row_idx'], 2, "success")
                        print(f"       Marked 'success' in Cloud Sheet")
                    except:
                        print(f"       Failed to update Cloud Sheet status")
                time.sleep(1)
            print("\n Batch Processing Complete")
    else:
        print("\n💡 No pending tasks found.")

    input("\nPress Enter to exit...")
