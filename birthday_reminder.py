import pandas as pd
import datetime
from plyer import notification
import pyttsx3
import sys

# ✅ Set up logging
log_file = open("log.txt", "w")
sys.stdout = log_file
sys.stderr = log_file

# ✅ Initialize speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    print(f"Saying: {audio}")
    engine.say(audio)
    engine.runAndWait()

def notification1(title, msg):
    print(f"Showing notification: {title} - {msg}")
    notification.notify(
        title=title,
        message=msg,
        app_icon="birthday_icon.ico",  # optional, put your .ico file here
        timeout=5
    )

# ✅ Load Excel
try:
    df = pd.read_excel("birthday_dates.xlsx")
    print("Excel loaded successfully.")
except Exception as e:
    print("Failed to load Excel:", e)
    sys.exit()

# ✅ Get today's date
today = datetime.datetime.now().strftime("%d-%m")
print("Today's date:", today)

# ✅ Check each row
found = False
for index, item in df.iterrows():
    name = str(item["name"]).strip()
    bd = str(item["birthday"]).strip()
    print(f"Checking: {name} - {bd}")

    if today == bd:
        found = True
        message = f"It's {name}'s birthday today!"
        notification1(" Birthday Alert ", message)
        speak(message)
        break

if not found:
    print("No birthdays today.")

# ✅ Done
print("Script finished running successfully.")

log_file.close()
