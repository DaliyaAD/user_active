# -*- coding: utf-8 -*-
"""
Created on Sun Apr  9 14:06:36 2023

@author: daliya
"""

import win32gui
import win32api 
import win32con
import win32process
import time
import winerror
#Change these values to match the title of the remote desktop window
number1 = input ("Enter X number tcremote(X): ")
number2 = input("Enter number XXX REE-XXX: ")
WINDOW_TITLE = f"tcremote ({number1}) - REEE-{number2} - Remote Desktop Connection"
WINDOW_CLASS = "TscShellContainerClass"

#Get the window handle for the remote desktop window
def get_window_handle():
    hwnd = win32gui.FindWindow(WINDOW_CLASS, WINDOW_TITLE)
    if hwnd == 0:
        raise Exception("Window not found")
    return hwnd

#Bring the window to the foreground
def set_window_pos(hwnd):
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
    if win32gui.GetForegroundWindow() == hwnd:
        return
    win32gui.BringWindowToTop(hwnd)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
    
#Simulate mouse click on the window
def simulate_mouse_click(hwnd):
    x, y = win32gui.ClientToScreen(hwnd, (0,0))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 10, 10)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 50, 50)

#Simulate mouse movement on the window
def simulate_mouse_movement(hwnd):
    x, y = win32gui.ClientToScreen(hwnd, (0,0))
    win32api.SetCursorPos((x+30, y+30))    
    print("done")
    
#Minimize remote window     
def minimize_window(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    
#Move cursor back    
def return_cursor(x,y):
    win32api.SetCursorPos((x,y))
    
    
#Main loop
def main():
    hwnd = get_window_handle()   
    
    while True:
        x, y = win32gui.GetCursorPos()
        
        set_window_pos(hwnd)
        simulate_mouse_movement(hwnd)
        simulate_mouse_click(hwnd)
        minimize_window(hwnd)
        return_cursor(x,y)
        time.sleep(28*60) #Wait for 5 mintes before repeating
        
        
if __name__ == "__main__":
     main()
        
