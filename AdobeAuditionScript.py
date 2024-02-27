import subprocess
import time
import pyautogui

from Constants import ADOBE_AUDITION_PATH,_multiTrack
from SharedGlobalVariables import Singleton




def open_adobe_audition():
    # Open Adobe Audition 
    adobe_audition_path = ADOBE_AUDITION_PATH
    subprocess.Popen(adobe_audition_path)
    # time.sleep(10)

def startRecording():
    pyautogui.hotkey('shift', 'space') 

def stopRecording():
    pyautogui.hotkey('space') 

def startRecordingSide():
    # Define the target position
    target_x, target_y = 738, 767

    # Move the cursor to the target position
    pyautogui.moveTo(target_x, target_y)

    # Click at the target position
    pyautogui.click()

def stopRecordingSide():
    # Define the target position
    target_x, target_y = 532, 770

    # Move the cursor to the target position
    pyautogui.moveTo(target_x, target_y)

    # Click at the target position
    pyautogui.click()

    # Save File
    # Press and hold down the 'ctrl', 'alt', and 'shift' keys
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('alt')
    pyautogui.keyDown('shift')

    # Press the 's' key
    pyautogui.press('s')

    # Release all the keys
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('alt')
    pyautogui.keyUp('shift')

    pyautogui.press('enter')
   
def control_adobe_audition(statement_id):
    # Automation for Adobe Audition
    # create recording window 
    pyautogui.hotkey('ctrl','shift', 'n')

    # 
    filename = "audio_" + Singleton.username + "_" + str(statement_id)
    
    pyautogui.typewrite(filename)
    pyautogui.press('enter')

    Singleton.activeFilename = filename
    startRecording()
    # Wait for user to press spacebar again to stop recording
    # input("Press spacebar to stop recording and save audio...")
    # print("Recording stopped and saved.")

def control_adobe_audition_side(statement_id):
    # choose recording window 
    multi_track(statement_id)

def createMulti(statement_id):
    pyautogui.hotkey('ctrl', 'n')
   
    filename = "audio_" + Singleton.username + "_" + str(statement_id)
    
    pyautogui.typewrite(filename)
    pyautogui.press('enter')
    
    # Define the target position
    target_x, target_y = 534, 221

    # Move the cursor to the target position
    pyautogui.moveTo(target_x, target_y)

    # Click at the target position
    pyautogui.click()
    
    Singleton.activeFilename = filename
   
    startRecordingSide()

def createSingle(statement_id):
    pyautogui.hotkey('ctrl','shift', 'n')
    filename = "audio_" + Singleton.username + "_" + str(statement_id)
    pyautogui.typewrite(filename)
    pyautogui.press('enter')
    Singleton.activeFilename = filename
    startRecording()

def multi_track(statement_id):
    if _multiTrack :
        createMulti(statement_id)   
    else:
        createSingle(statement_id)
