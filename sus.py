import sys
import tkinter as tk
from PIL import Image, ImageTk
import random
import threading
import time
import pygame
import os
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pynput.keyboard import Listener, Key

stop_flag = False
windows = []  # Keep track of opened Tkinter windows

def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def set_volume_max_unmute():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(1.0, None)  # Max volume
    volume.SetMute(0, None)  # Unmute

def play_sound_background(sound_file_path):
    pygame.mixer.init()
    pygame.mixer.music.load(sound_file_path)
    pygame.mixer.music.play(-1)  # Loop indefinitely
    global stop_flag
    while not stop_flag:
        time.sleep(0.1)
    pygame.mixer.music.stop()  # Stop the music when stop_flag is True

def display_random_image(image_path):
    root = tk.Tk()
    root.overrideredirect(True)  # No window border
    screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
    img = Image.open(image_path)
    imgTk = ImageTk.PhotoImage(img, master=root)
    label = tk.Label(root, image=imgTk)
    label.pack()
    x_position, y_position = random.randint(0, screen_width - img.width), random.randint(0, screen_height - img.height)
    root.geometry(f'+{x_position}+{y_position}')
    root.protocol("WM_DELETE_WINDOW", lambda: close_window(root))  # Handle window close event
    windows.append(root)  # Add window to list
    root.mainloop()

def close_window(window):
    window.destroy()
    windows.remove(window)

def close_windows():
    while windows:  # Ensure all windows are closed
        window = windows.pop()
        window.destroy()

def continuous_display(interval, repetitions, image_path):
    for _ in range(repetitions):
        if stop_flag:  # Exit early if stop_flag is True
            close_windows()
            return
        threading.Thread(target=display_random_image, args=(image_path,), daemon=True).start()
        time.sleep(interval)

def on_press(key):
    global stop_flag
    if key == Key.char('q'):
        stop_flag = True
        close_windows()  # Close all windows when 'q' is pressed
        sys.exit()  # Exit the program

listener = Listener(on_press=on_press, daemon=True)
listener.start()

if __name__ == "__main__":
    sound_file_path = resource_path("sus.mp3")
    image_file_path = resource_path("sus.png")
    
    # Starting the background music in a daemon thread ensures it doesn't prevent program termination.
    threading.Thread(target=play_sound_background, args=(sound_file_path,), daemon=True).start()

    display_interval = 0.01
    display_repetitions = 1
    
    duration = 60  # seconds
    timer = threading.Timer(duration, lambda: stop_flag == True)  # Set timer to stop the program after the specified duration
    timer.start()

    continuous_display(display_interval, display_repetitions, image_file_path)
    set_volume_max_unmute()
    
    stop_flag = True  # Ensure flag is set to stop threads
    close_windows()  # Close all Tkinter windows
    sys.exit(0)  # Exit the program
