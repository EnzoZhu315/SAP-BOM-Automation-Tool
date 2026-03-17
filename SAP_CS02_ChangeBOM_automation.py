import win32com.client
import subprocess
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import ssl

# ==========================================
# 1. Configuration (Masked Information)
# ==========================================
# [DE-IDENTIFIED] Placeholders for corporate credentials
USER_INFO = {"user": "YOUR_USERNAME", "password": "YOUR_PASSWORD"}
SAP_SYSTEM_ID = "YOUR_SAP_SYSTEM_CONNECTION_NAME"
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# Google Sheet Configuration
GS_KEY = 'YOUR_GOOGLE_SHEET_ID_HERE'
CREDENTIALS_JSON = "client_secret.json"

try:
    ssl._create_default_https_context = ssl._create_unverified_context
except:
    pass


# ==========================================
# 2. Fetch Tasks from Google Sheet
# ==========================================
def get_cs02_tasks():
    print("\n --- Step 1: Fetching Tasks from Google Sheet ---")
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_JSON, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(GS_KEY)
        
        # Access the specific worksheet for CS02
        sheet = spreadsheet.worksheet("CS02_Task_List")

        all_data = sheet.get_all_values()
        tasks = []
        for i, row in enumerate(all_data):
            if i == 0: continue # Skip header
            
            p_no = str(row[0]).strip()
            status = row[1] if len(row) > 1 else ""
            
            # Filter logic: Material starts with 'P' and is not yet processed
            if p_no.upper().startswith('P') and status.lower() != "success":
                tasks.append({"p_no": p_no, "row_idx": i + 1})

        print(f" Connected to Sheet. Tasks found: {len(tasks)}")
        return tasks, sheet
    except Exception as e:
        print(f" Google Sheet Access Failed: {e}")
        return [], None


# ==========================================
# 3. SAP Logon Logic
# ==========================================
def get_sap_session():
    print("\n --- Step 2: Logging into SAP ---")
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

        if session.findById("wnd[0]/usr/txtRSYST-BNAME", False):
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USER_INFO['user']
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = USER_INFO['password']
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
            # Handle multi-logon popups
            try:
                if "wnd[1]" in str(session.ActiveWindow.Name):
                    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                pass
        print(" SAP Logon Successful!")
        return session
    except Exception as e:
        print(f" SAP Connection Failed: {e}")
        return None


# ==========================================
# 4. CS02 Transaction Logic
# ==========================================
def run_sap_cs02(session, p_number):
    print(f"\n Maintaining Material: {p_number}")
    try:
        # Navigate to CS02 (Change BOM)
        session.findById("wnd[0]").resizeWorkingPane(88, 29, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS02"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)

        # Initial Screen Input
        # [DE-IDENTIFIED] Masked Plant Code and Change Number (AENNR)
        session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = p_number
        session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "YOUR_PLANT_CODE"
        session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
        session.findById("wnd[0]/usr/ctxtRC29N-AENNR").text = "YOUR_CHANGE_NUMBER"

        #  Trigger initial data submission
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        #  Exception Handling: Bypass initial system warnings (e.g., date alerts)
        print("      ⚠️ Bypassing initial system warnings...")
        for _ in range(2):
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)

        # Table Control Path
        table_path = "wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT"

        # Check for specific component slot (e.g., Row Index 3)
        try:
            current_comp = session.findById(f"{table_path}/ctxtRC29P-IDNRK[2,3]").text.strip()
        except:
            current_comp = ""

        if current_comp == "":
            print(f"       Appending new component...")
            session.findById(f"{table_path}/ctxtRC29P-POSTP[1,3]").text = "L"
            session.findById(f"{table_path}/ctxtRC29P-IDNRK[2,3]").text = "NEW_COMPONENT_ID"
            session.findById(f"{table_path}/txtRC29P-MENGE[4,3]").text = "0.5"
            session.findById(f"{table_path}/ctxtRC29P-MEINS[5,3]").text = "PC"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)
        else:
            print(f"      Component already exists: {current_comp}")

        # ---  Loop Enter Key to clear Material Status Warnings (e.g., Status C9) ---
        print("       Clearing item-level status warnings...")
        for _ in range(3):
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)

        # Save Transaction
        print("       Committing changes to SAP...")
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(2)

        # Validation via Status Bar (SBar)
        status_msg = session.findById("wnd[0]/sbar").text
        print(f"       SAP Feedback: {status_msg}")

        success_keywords = ["changed", "created", "no changes"]
        if any(kw in status_msg.lower() for kw in success_keywords):
            print(f"      Transaction Successful!")
            return True
        else:
            print(f"       Save failed. Check SAP status bar.")
            return False

    except Exception as e:
        print(f"       Script Exception: {e}")
        return False


# ==========================================
# 5. Main Execution
# ==========================================
if __name__ == "__main__":
    tasks, sheet_obj = get_cs02_tasks()
    if tasks and sheet_obj:
        sap_session = get_sap_session()
        if sap_session:
            for item in tasks:
                if run_sap_cs02(sap_session, item['p_no']):
                    try:
                        # Update Cloud Sheet status upon success
                        sheet_obj.update_cell(item['row_idx'], 2, "success")
                        print(f"       Cloud Sheet updated.")
                    except:
                        print(f"       Cloud Sheet update failed.")
                time.sleep(1)
            print("\n Batch Processing Finished.")
    else:
        print("\n Info: No pending P-Numbers found.")

    input("\nPress Enter to exit...")
