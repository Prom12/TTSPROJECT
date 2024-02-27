import pyautogui
import keyboard

dev_mode = False
_multiTrack = True
side_by_side = True
_getMouseLocation = False
 
test_path = "C:\\Users\\Prom\\Desktop\\TTS PROJECT\\"
dev_path = "C:\\Users\\isaac\\Desktop\\TTS PROJECT\\"

if dev_mode:
    # All File Paths
    TTS_CONTENT_PATH = "C:\\Users\\isaac\\Desktop\\TTS PROJECT\\TTSCONTENT.xlsx"
    TTS_DISPLAY_PATH = "C:\\Users\\isaac\\Desktop\\TTS PROJECT\\TTSDISPLAY.docx"
    TTS_RECORDERS_PATH = "C:\\Users\\isaac\\Desktop\\TTS PROJECT\\TTSRECORDERS.xlsx"

    # APPLICATION PATHS
    ADOBE_AUDITION_PATH = "C:\\Program Files\\Adobe\\Adobe Audition 2024\\Adobe Audition.exe"
    MAIN_WORD_PATH = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
    WORD_PATH =  "C:\\Users\\isaac\\Desktop\\TTS PROJECT\\TTSDISPLAY.docx"
else:
    # All File Paths
    TTS_CONTENT_PATH = test_path + "TTSCONTENT.xlsx"
    TTS_DISPLAY_PATH = test_path + "TTSDISPLAY.docx"
    TTS_RECORDERS_PATH = test_path + "TTSRECORDERS.xlsx"

    # APPLICATION PATHS
    ADOBE_AUDITION_PATH = "C:\\Program Files\\Adobe\\Adobe Audition 2024\\Adobe Audition.exe"
    MAIN_WORD_PATH = "C:\\Program Files\\Microsoft Office\\Office16\\WINWORD.EXE"
    WORD_PATH = test_path + "TTSDISPLAY.docx"

def SIDE_BY_SIDE(functionState):
    if side_by_side:
        return functionState
    else:
        return

def GET_MOUSE_LOCATION():
    if _getMouseLocation :
        while True:
            # Get the current cursor position
            x, y = pyautogui.position()

            # Print the cursor position
            print(f"Cursor position (x, y): {x}, {y}")
            
            # Check if spacebar is pressed
            if keyboard.is_pressed('space'):
                print("Spacebar clicked. Exiting...")
                break
       
        exit()