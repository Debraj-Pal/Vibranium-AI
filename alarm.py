import win32com.client
import os
import datetime
import time

# convert text into speech
def speak(text):
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.Rate = 1
        speaker.Volume = 100

        speaker.Speak(text)
        time.sleep(0.5)  # Adding delay to ensure audio resources are released
    except Exception as e:
        print(f"Audio Error: {e}")

extractedtime=open("Alarm.txt","rt")
time_data = extractedtime.read()
extractedtime.close()

deletetime= open("Alarm.txt","r+")
deletetime.truncate(0)
deletetime.close()

def ring(time_str):
    timeset = str(time_str)
    timenow = timeset.replace("Vibranium","")
    timenow = timenow.replace("set an alarm","")
    timenow = timenow.replace(" and ",":").strip()
    timenow = timenow.replace("for", "").replace("at", "").strip()

    if ":" in timenow:
        time_parts = timenow.split(":")
        hour = time_parts[0].strip().zfill(2)
        minute = time_parts[1].strip().zfill(2)
        second = time_parts[2].strip().zfill(2) if len(time_parts) == 3 else "00"
        timenow = f"{hour}:{minute}:{second}"
        
    print(f"Vibranium: Alarm set for: {timenow}")
    speak(f"Alarm set for: {timenow}")

    while True:
        currenttime = datetime.datetime.now().strftime("%H:%M:%S")

        if currenttime == timenow:
            print("Vibranium: Alarm ringing, sir")
            speak("Alarm ringing, sir")
            music_path = r"F:\Software Back up\Entertainment\Songs\Music\Bones(PaglaSongs).mp3"
            if os.path.exists(music_path):
                os.startfile(music_path)
            else:
                print(f"Audio file not found at: {music_path}")
            time.sleep(60)
            print("Alarm execution finished. Shutting down script cleanly.")
            break  # Exits the loop and terminates the file so it doesn't run forever in the background
            
        time.sleep(1)  

ring(time_data)
