import time
import win32com.client

# Create a Shell object
shell = win32com.client.Dispatch("WScript.Shell")

# Open Chrome
shell.Run("chrome.exe")

# Wait for Chrome to open
time.sleep(2)

# Use the Shell object to simulate the keyboard
shell.AppActivate("Chrome")  # Focus Chrome window, needs exact window title if multiple windows are open

# Open a new tab with the desired YouTube video
shell.SendKeys("https://www.youtube.com/watch?v=xAQ7pemTSLE")  # Replace VIDEO_ID with the actual ID
shell.SendKeys("{ENTER}")

# Wait for 5 seconds while the video plays
time.sleep(12)

# Close the current tab using Ctrl+F4
shell.SendKeys("^%{F4}")

# Close Chrome after a short delay if needed
time.sleep(1)
shell.SendKeys("%{F4}")
