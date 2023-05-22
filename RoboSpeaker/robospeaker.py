import win32com.client
import os
speaker = win32com.client.Dispatch("SAPI.SpVoice")

while(1):
    text = input("Enter what you want to speak:").upper()
    if text == "Q":
        speaker.speak("Thank you for using RoboSpeaker.")
        break
    speaker.speak(text)
