import time
import pyautogui
import keyboard
import pygetwindow as gw
import pyautogui

def minimize_all():
    pyautogui.hotkey('win','d')

def bring_to_foreground_and_fullscreen(window_title):
    try:
        minimize_all()

        audition_window = gw.getWindowsWithTitle(window_title)

        if audition_window:
            audition_window[0].activate()
            audition_window[0].maximize()
        # time.sleep(0.2)
    except Exception as e:
        print(str(e))

def make_content_active(window_title):
    try:
        # minimize_all()
        audition_window = gw.getWindowsWithTitle(window_title)

        if audition_window:
            audition_window[0].activate()
    except Exception as e:
        print(str(e))

def wait_for_spacebar():
    while True:
        if keyboard.is_pressed('space'):
            break
        time.sleep(0.1)

# Function to get screen dimensions
def get_screen_dimensions():
    screen_width, screen_height = pyautogui.size()  # Get primary screen
    return screen_width, screen_height

# Function to position windows side by side on the screen
def position_windows_side_by_side():
    # minimize all first
    # minimize_all()
    screen_width, screen_height = get_screen_dimensions()
    # Define window titles, dimensions, and positions
    window_titles = ["Adobe Audition", "TTSDISPLAY"]

    # Calculate window width and position
    window_width = screen_width // len(window_titles)
    window_positions = [(i * window_width, 0) for i in range(len(window_titles))]
    # window_width = screen_width // 2
    window_height = screen_height
    # window_positions = [(0, 0), (window_width // 2, 0)]
    
    for title, position in zip(window_titles, window_positions):
        window = gw.getWindowsWithTitle(title)
        if window:
            window = window[0]
            window.activate()  # Bring the window into focus
            # window.move(0,0)
            window_width = min(screen_width // 2,window.width)
            window.resizeTo(window_width, window_height)
            window.move(*position)
            print(f"{title} window resized and moved.")
        else:
            print(f"{title} window not found.")
