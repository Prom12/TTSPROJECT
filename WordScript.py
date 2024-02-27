import subprocess
import time
import win32com.client

from Constants import TTS_RECORDERS_PATH,WORD_PATH,MAIN_WORD_PATH
from SharedGlobalVariables import Singleton
from SharedFunctions import bring_to_foreground_and_fullscreen,wait_for_spacebar,make_content_active
from AdobeAuditionScript import stopRecording,control_adobe_audition,control_adobe_audition_side,stopRecordingSide


def open_word():
    # Path to the Word document
    word_path = WORD_PATH
    try:
        subprocess.Popen([MAIN_WORD_PATH, word_path])
        print("Microsoft Word document opened successfully.")
    except Exception as e:
        print("Error:", e)

def update_word_document(texts):
    try:
        word_app = win32com.client.GetActiveObject("Word.Application")
    except Exception as e:
        if isinstance(e, OSError) and e.winerror == -2147221021:
            print("Microsoft Word is not running or not accessible.")
        else:
            print(f"Error: {e}")
        # try:
        #     word_app = win32com.client.Dispatch("Word.Application")
        # except Exception as e:
        #     print(f"Error: Unable to create or connect to Microsoft Word instance: {e}")
        #     return update_word_document(texts)
    
    word_app.Visible = True

    time.sleep(1)
    
    word_doc = word_app.ActiveDocument

    # Check if the document is untitled
    if word_doc.FullName == "":
        print("Warning: Untitled document detected. Please ensure Word has an active document.")


    
    word_doc.Content.Delete()
    
    
    for text in texts:
        if text != texts[-1]:
            text +="\n\n"
        word_doc.Range().InsertAfter(text)

    word_doc.Save()

def control_word(Statement,statement_id):
    # global activeFilename, activeWB, activeSheet
    texts= [
        Statement, 
        ".......................................................", 
        "Press spacebar to start Recording"
    ]

    # Automation for Microsoft Word
    
    bring_to_foreground_and_fullscreen("TTSDISPLAY")
    update_word_document(texts)
    wait_for_spacebar()

    bring_to_foreground_and_fullscreen("Adobe Audition")
    control_adobe_audition(statement_id)
    texts[2] = "Recording Started ............."

    bring_to_foreground_and_fullscreen("TTSDISPLAY")
    update_word_document(texts)
    wait_for_spacebar()
    
    bring_to_foreground_and_fullscreen("Adobe Audition")
    stopRecording()
    texts[2] = "Recording Stopped"

    bring_to_foreground_and_fullscreen("TTSDISPLAY")    
    update_word_document(texts)
    wait_for_spacebar()
    
    # Save Recording Information to Workbook
    max_row = Singleton.activeSheet.max_row
    Singleton.activeSheet.cell(row=max_row + 1, column=1, value= Singleton.username)  # Username
    Singleton.activeSheet.cell(row=max_row + 1, column=2, value= Singleton.activeFilename)  # Audio Name
    Singleton.activeSheet.cell(row=max_row + 1, column=3, value= Statement)  # Statement (can be filled later)
    Singleton.activeSheet.cell(row=max_row + 1, column=4, value= statement_id)  # StatementID
    Singleton.activeWB.save(TTS_RECORDERS_PATH)
    # pyautogui.click(100,100)
    # pyautogui.hotkey('ctrl','a')
    # pyautogui.hotkey('delete')
    # pyperclip.copy(Statement)
    # time.sleep(1)
    # pyautogui.hotkey('ctrl','v')

    # print("\n\n")
    
    
    # Wait for user confirmation to proceed

def control_word_side(Statement,statement_id):
    texts= [
        Statement, 
        ".......................................................", 
        "Press spacebar to start Recording"
    ]

    # Automation for Microsoft Word
    make_content_active("TTSDISPLAY")
    update_word_document(texts)
    wait_for_spacebar()

    make_content_active("Adobe Audition")
    control_adobe_audition_side(statement_id)
    # time.sleep(0.2)
    
    texts[2] = "Recording Started ............."

    # time.sleep(0.2)

    make_content_active("TTSDISPLAY") # check this line
    update_word_document(texts) # and this
    wait_for_spacebar()
     
    make_content_active("Adobe Audition")
    stopRecordingSide()
    texts[2] = "Recording Stopped"

    make_content_active("TTSDISPLAY")    
    update_word_document(texts)
    wait_for_spacebar()
    
    # Save Recording Information to Workbook
    max_row = Singleton.activeSheet.max_row
    Singleton.activeSheet.cell(row=max_row + 1, column=1, value= Singleton.username)  # Username
    Singleton.activeSheet.cell(row=max_row + 1, column=2, value= Singleton.activeFilename)  # Audio Name
    Singleton.activeSheet.cell(row=max_row + 1, column=3, value= Statement)  # Statement (can be filled later)
    Singleton.activeSheet.cell(row=max_row + 1, column=4, value= statement_id)  # StatementID
    Singleton.activeWB.save(TTS_RECORDERS_PATH)