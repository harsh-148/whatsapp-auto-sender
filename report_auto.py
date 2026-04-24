import win32com.client as win32
import pyautogui
import time
import os

# --- CONFIG ---
RECIPIENTS = [
    (" 87########", "M2:T24", "S1"),
    (" 98########", "C2:J24", "S2"),
    (" 98########", "m2:t24", "S2"),
    (" 98########", "C2:R24", "S3"),
    (" 99########", "C2:R24", "S4"),
    (" 98#######",  "C2:J24", "S5"),
    (" 98########", "C2:J24", "S6"),
    (" 96########", "C2:J24", "S7"),
    (" 93########", "C2:J24", "S8"),
    (" 98########", "B2:H24", "S9"),
]

# ✅ Use forward slashes to avoid unicode escape issues
EXCEL_PATH = "Path/To/Your/Report.xlsx"


def run_perfect_send():
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    
    try:
        wb = excel.Workbooks.Open(EXCEL_PATH)
        
        for phone, cell_range, sheet_name in RECIPIENTS:
            print(f"📸 Step 1: Copying {sheet_name} range {cell_range}...")
            ws = wb.Worksheets(sheet_name)
            ws.Activate()
            
            excel.Range(cell_range).CopyPicture(Appearance=1, Format=2)
            time.sleep(2)

            print(f"📱 Step 2: Switching to WhatsApp for {phone}...")
            os.startfile(f"whatsapp://send?phone={phone}")
            time.sleep(6)

            screen_width, screen_height = pyautogui.size()
            pyautogui.click(screen_width // 2, screen_height // 2)
            time.sleep(1)

            print("📋 Step 3: Pasting and Sending...")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(5)
            
            pyautogui.press('enter')
            print(f"✅ Sent to {phone}")
            time.sleep(3)

        wb.Close(False)

    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        excel.Quit()


# ✅ Fixed: double underscores
if __name__ == "__main__":
    run_perfect_send()