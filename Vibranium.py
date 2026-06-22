import win32com.client
import time
import speedtest
import speech_recognition as sr
import pyaudio
import datetime
import os
import webbrowser
import pywhatkit as kit
from config import apikey
from config import api_key
from config import weather_api_key
from google import genai
from google.genai import types
import requests
from bs4 import BeautifulSoup
import pyautogui
import cv2
import pyjokes
from datetime import timedelta
import sys
import ollama
import json
import threading
import subprocess

# convert our voice to speech
def takecommand():
    r = sr.Recognizer()
    r.pause_threshold = 1.0
    r.dynamic_energy_threshold = True

    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source, duration=0.5)

        try:
            # Added a timeout and phrase_time_limit so it doesn't get stuck
            audio = r.listen(source, timeout=60, phrase_time_limit=10)
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query.lower()

        except sr.WaitTimeoutError:
            print("Microphone timed out. Resetting...", end="\r")
            return "none"
        except Exception as e:
            print("...Listening...", end="\r")
            return "none"

# convert text into speech
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 1
speaker.Volume = 100
def speak(text):
    try:
        speaker.Speak(text)
        time.sleep(0.5)  # Adding delay to ensure audio resources are released
    except Exception as e:
        print(f"Audio Error: {e}")

# to make Vibranium wish you
def wish():
    hour = int(datetime.datetime.now().hour)
    time_str = datetime.datetime.now().strftime("%#I:%M %p")
    
    if (hour >= 6) and (hour < 12):
        greeting = f"Good Morning sir, it is {time_str}"
    elif (hour >= 12) and (hour < 17):
        greeting = f"Good Afternoon sir, it is {time_str}"
    elif (hour >= 17) and (hour < 20):
        greeting = f"Good Evening sir, it is {time_str}"
    elif (hour >=20) and (hour < 24):
        greeting = f"Good Night sir, it is {time_str}"
    else:
        greeting = f"Good Night sir, it is {time_str}"
    full_message = f"{greeting}. I am Vibranium, your personal AI assistant. How can I help you today?"

    print(f"Vibranium: {full_message}")
    speak(full_message)

#history file to save the conversation history for context in Vibranium's responses. This allows the assistant to maintain context across interactions and provide more relevant answers based on previous messages.
HISTORY_FILE = "history.json"

def load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    with open(HISTORY_FILE, "r") as f:
        try:
            return json.load(f)
        except:
            return []

def save_message(role, content):
    try:
        history = load_history()
        # Add new message
        history.append({"role": role, "parts": [{"text": content}]})
        # Save back to file
        with open(HISTORY_FILE, "w") as f:
           json.dump(history, f)

    except Exception as e:
        print(f"Could not save to history: {e}")     

#google api key for genai client

GOOGLE_API_KEY=apikey
client = genai.Client(api_key=GOOGLE_API_KEY)

config = types.GenerateContentConfig(
    temperature=0.9,
    top_p=1.0,
    top_k=1,
    max_output_tokens=2048,
    system_instruction="You are a AI voice assistant named Vibranium built by Debraj Paul. Provide clear, conversational answers. Do not use markdown formatting, bullet points, headers, emojis, or complex tables. Write your response entirely in plain text sentences so it can be read aloud easily.",
    safety_settings=[
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_HARASSMENT,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
    ]
)

initial_history = load_history()

chat_session = client.chats.create(
    model="gemini-3.5-flash",
    config=config,
    history=initial_history
)

# Function to reset memory by clearing the history file and resetting the chat session with Gemini. 
# This allows Vibranium to start fresh without any previous context, which can be useful if the assistant is providing irrelevant answers or if the user wants to clear the conversation history for privacy reasons.
def reset_memory():
    global chat_session # This allows us to modify the global session
    
    # 1. Delete the history file to clear past conversations
    if os.path.exists(HISTORY_FILE):
        os.remove(HISTORY_FILE)
    
    # 2. Reset the active RAM memory
    chat_session = client.chats.create(
        model="gemini-3.5-flash",
        config=config
    )
    print("Memory cleared, sir. I have started fresh.")
    speak("Memory cleared, sir. I have started fresh.")

OLLAMA_MODE = False

def chat(query):
    global OLLAMA_MODE 
    
    reply = "" 
    system_instruction = (
        "You are an AI voice assistant named Vibranium built by Debraj Paul. "
        "Provide clear, conversational answers. Do not use markdown formatting, "
        "bullet points, headers, emojis, or complex tables. Write your response "
        "entirely in plain text sentences so it can be read aloud easily."
    )
    # --- PATH A: THE SWITCH IS ON ---
    if OLLAMA_MODE:
        try:
            raw_history = load_history()
            ollama_messages = [{"role": "system", "content": system_instruction}]
            
            for msg in raw_history:
                ollama_role = "user" if msg.get("role") == "user" else "assistant"
                try:
                    content = msg["parts"][0]["text"]
                    ollama_messages.append({"role": ollama_role, "content": content})
                except (KeyError, IndexError):
                    continue
                    
            ollama_messages.append({"role": "user", "content": query})
            response = ollama.chat(model='llama3.2:1b', messages=ollama_messages)
            reply = response['message']['content']
            
            # Save to history
            save_message("user", query)
            save_message("model", reply)
            
        except Exception as e_local:
            print(f"Ollama Error: {e_local}")
            speak("I am sorry, both Gemini and my local assistant are unavailable.")
            return

    # --- PATH B: THE SWITCH IS OFF (Vibranium Will try Gemini normally) ---
    else:
        try:
            response = chat_session.send_message(query)
            reply = response.text
            
            # --- SAVE TO MEMORY ---
            save_message("user", query)
            save_message("model", reply)

        except Exception as e:
            if "429" in str(e):
                print("Gemini quota hit! Switching to Local AI (Ollama)...")
                speak("Sorry sir, I have hit my Gemini API quota. Switching to my local assistant, please wait.")
                OLLAMA_MODE = True 
                
                try:
                    print("Ollama is assembling conversation context... Please hold on...")
                    raw_history = load_history()
                    ollama_messages = [{"role": "system", "content": system_instruction}]
                    
                    for msg in raw_history:
                        ollama_role = "user" if msg.get("role") == "user" else "assistant"
                        try:
                            content = msg["parts"][0]["text"]
                            ollama_messages.append({"role": ollama_role, "content": content})
                        except (KeyError, IndexError):
                            continue
                            
                    ollama_messages.append({"role": "user", "content": query})
                    response = ollama.chat(model='llama3.2:1b', messages=ollama_messages)
                    reply = response['message']['content']
                    
                    # --- TO SAVE LOCAL CONVERSATION TO MEMORY TOO ---
                    save_message("user", query)
                    save_message("model", reply)
                    
                except Exception as e_local:
                    print(f"Ollama Error: {e_local}")
                    speak("I am sorry, both Gemini and my local assistant are unavailable.")
                    return
            else:
                print(f"Gemini Error: {e}")
                speak("Sorry, I encountered a connection error.")
                return

    print(f"Vibranium: {reply}")

    clean_reply = reply.replace("**", "").replace("###", "").replace("##", "").replace("#", "")
    clean_reply = clean_reply.replace("`", "").replace("*", "").replace("- ", "").replace("|", " ")
    clean_reply = clean_reply.encode('ascii', 'ignore').decode('ascii')
    clean_reply = clean_reply.replace('\n', ' ').replace('\r', ' ')

    if clean_reply.strip():
        speak(clean_reply)

def get_alarm_input():
    result = [None]
    event = threading.Event()

    def from_text():
        try:
            text = input("Type the time (e.g. 10 and 30 and 00): ").strip()
            if text and not event.is_set():
                result[0] = text
                event.set()
        except:
            pass

    def from_voice():
        try:
            voice = takecommand()
            if voice != "none" and not event.is_set():
                result[0] = voice
                event.set()
        except:
            pass

    t1 = threading.Thread(target=from_text, daemon=True)
    t2 = threading.Thread(target=from_voice, daemon=True)

    t1.start()
    t2.start()

    event.wait()  # Blocks until either text or voice is entered
    return result[0]

def alarm(query):
    timehere=open("Alarm.txt","w")
    timehere.write(query)
    timehere.close()
    subprocess.Popen(["python", "alarm.py"])

def news():
    # Dynamically get a date from 7 days ago to stay safely within the 30-day free limit
    from datetime import date, timedelta
    safe_date = (date.today() - timedelta(days=7)).isoformat()
    
    # Using everything endpoint to search for India news safely
    newsurl = f"https://newsapi.org/v2/everything?q=India&from={safe_date}&sortBy=popularity&language=en&apiKey={api_key}"
    
    try:
        response = requests.get(newsurl)
        main_page = response.json()
        
        if main_page.get("status") == "error":
            error_msg = main_page.get("message", "Unknown API error.")
            print(f"NewsAPI Error: {error_msg}")
            speak("Sir, the news API returned an error. Please check your console log.")
            return

        articles = main_page.get("articles", [])
        
        if not articles:
            print("Vibranium: No articles found matching your request.")
            speak("No articles found matching your request, sir.")
            return

        # We will now store dictionaries containing both titles and descriptions
        news_stories = []
        day = ["first", "second", "third", "fourth", "fifth"] # Capped at 5 so it doesn't talk too long
        
        for ar in articles:
            title = ar.get("title")
            description = ar.get("description")
            
            # Avoid blank titles, blank summaries, or removed articles
            if title and description and "[Removed]" not in title:
                # Clean up any weird character artifacts (like the El Niño question mark bug)
                clean_title = title.encode('ascii', 'ignore').decode('ascii')
                clean_desc = description.encode('ascii', 'ignore').decode('ascii')
                
                news_stories.append({
                    "title": clean_title,
                    "description": clean_desc
                })
            
        total_news_to_read = min(len(day), len(news_stories))
        
        if total_news_to_read == 0:
            print("Vibranium: No valid articles available right now.")
            speak("No valid articles available right now, sir.")
            return

        for i in range(total_news_to_read):
            story_title = news_stories[i]["title"]
            story_desc = news_stories[i]["description"]
            
            # Print cleanly to the terminal console
            print(f"\n=================== {day[i].upper()} STORY ===================")
            print(f"Headline: {story_title}")
            print(f"Details : {story_desc}")
            
            # Vibranium announces the headline, then reads the descriptive paragraph summary
            speak(f"Today's {day[i]} news is: {story_title}. Here are the details. {story_desc}")
            
    except Exception as e:
        print(f"News Error: {e}")
        speak("Sorry sir, I was unable to retrieve the news at the moment.")

def message():
    print("Vibranium: Who do you want to message")
    speak("Who do you want to message")
    num=int(input(" "))

    print("Vibranium: What is the message sir")
    speak("What is the message sir")
    send_message = str(input(" "))
    scheduled_time = datetime.datetime.now() + timedelta(minutes=2)
    hour = scheduled_time.hour
    minute = scheduled_time.minute

    print(f"Vibranium: Scheduling message for {hour}:{minute:02d}...")
    kit.sendwhatmsg(f"+91{num}", send_message, hour, minute, wait_time=30, tab_close=True, close_time=3)
    print("Vibranium: Message delivered successfully sir!")
    speak("Message delivered successfully sir")

def load_sites():
    try:
        with open("sites.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        print("sites.json not found!")
        return []    

OWM_KEY = weather_api_key

if __name__ == '__main__':

    while True:
         query = takecommand()
         if "Yo Vibranium".lower() in query.lower() or "hello Vibranium".lower() in query.lower() or "hi Vibranium".lower() in query.lower() or "hey Vibranium".lower() in query.lower() or "Vibranium".lower() in query.lower():
            wish()

            while True:
                query=takecommand()

                if query == "none":
                    time.sleep(1)
                    continue
                    
                if "hello Vibranium".lower() in query.lower() or "hi Vibranium".lower() in query.lower() or "hey Vibranium".lower() in query.lower() or "Yo Vibranium".lower() in query.lower():
                    print("Vibranium: I am here sir, how can I help you")
                    speak("I am here sir, how can I help you")
                
                elif "clear memory".lower() in query.lower() or "reset memory".lower() in query.lower() or "forget everything".lower() in query.lower():
                    reset_memory()
                    continue 
                
                elif "open new tab".lower() in query.lower():
                    pyautogui.hotkey('ctrl', 't')
                    print("Vibranium: Opening new tab")
                    speak("Opening new tab")

                elif "show wi-fi speed".lower() in query.lower() or "check internet speed".lower() in query.lower() or "internet speed".lower() in query.lower() or "check wi-fi speed".lower() in query.lower():
                    try:
                       speak("Checking your internet speed, please wait...")
                       print("Connecting to the closest server...")
        
                       wifi = speedtest.Speedtest()
                       wifi.get_best_server()
                       print("Measuring speeds...")
                       download_speed = wifi.download() / 1_000_000
                       upload_speed = wifi.upload() / 1_000_000
                       upload_msg = f"Wi-fi upload speed is {upload_speed:.2f} Mbps"
                       download_msg = f"Wi-fi download speed is {download_speed:.2f} Mbps"        
                       print(upload_msg)
                       print(download_msg)
                       speak(upload_msg)
                       speak(download_msg)
        
                    except Exception as e:
                       print(f"Error checking speed: {e}")
                       speak("Sorry, I could not complete the speed test right now.")
        
                    time.sleep(2)


                elif "send message".lower() in query.lower():
                    message()

                elif "Search in Google".lower() in query.lower():
                    search_goog = query.lower().replace("search in google", "").replace("vibranium", "").strip()
                    if not search_goog:
                        print("Vibranium: What should I search on google")
                        speak("Sir, what should I search on google")
                        search_goog = takecommand().lower()
                    if search_goog and search_goog != "none":      
                        print(f"Vibranium: Searching {search_goog} sir")
                        speak(f"Searching {search_goog} sir")
                        webbrowser.open(f"https://www.google.com/search?q={search_goog.replace(' ', '+')}") 
                    else:
                        print("Vibranium: Sir, I didn't catch that. Search cancelled.")
                        speak("Sir, I didn't catch that. Search cancelled.")

                elif "Search this video in YouTube".lower() in query.lower():
                  print("Vibranium: What should I search on YouTube?")
                  speak("Sir, what should I search on YouTube?")
                  dm=takecommand().lower()
                  print(f"Vibranium: Searching {dm}")
                  speak(f"Searching {dm}")
                  kit.playonyt(dm)

                elif "search on youtube".lower() in query.lower() or "search in youtube".lower() in query.lower():
                   
                    search_topic = query.lower().replace("search on youtube", "").replace("vibranium", "").strip()
                    
                    if not search_topic:
                        print("Vibranium: What should I search for on YouTube?")
                        speak("Sir, what should I search for on YouTube?")
                        search_topic = takecommand().lower() 
                    
                    if search_topic and search_topic != "none":
                        print(f"Vibranium: Searching YouTube for '{search_topic}' sir...")
                        speak(f"Searching YouTube for {search_topic}, sir.")
                        formatted_url = f"https://www.youtube.com/results?search_query={search_topic.replace(' ', '+')}"
                        webbrowser.open(formatted_url)
                    else:
                        print("Vibranium: Sir, I didn't catch that. Search cancelled.")
                        speak("Sir, I didn't catch that. Search cancelled.")


                #YT video features
                elif "pause video".lower() in query.lower():
                    pyautogui.press("k")
                    print("Vibranium: Video paused")
                    speak("Video paused")
                elif "play video".lower() in query.lower():
                    pyautogui.press("k")
                    print("Vibranium: Playing video")
                    speak("Playing video")
                elif "mute video".lower() in query.lower():
                    pyautogui.press("m")
                    print("Vibranium: Video muted")
                    speak("Video muted")
                elif "unmute video".lower() in query.lower():
                    pyautogui.press("m")
                    print("Vibranium: Video unmuted")
                    speak("Video unmuted")
                elif "open cinema mode".lower() in query.lower():
                    pyautogui.press("t")
                    print("Vibranium: Cinema mode enabled")
                    speak("Cinema mode enabled")
                elif "exit cinema mode".lower() in query.lower():
                    pyautogui.press("t")
                    print("Vibranium: Cinema mode disabled")
                    speak("Cinema mode disabled")
                elif "full screen".lower() in query.lower():
                    pyautogui.press("f")
                    print("Vibranium: Full screen enabled")
                    speak("Full screen enabled")
                elif "exit full screen".lower() in query.lower():
                    pyautogui.press("f")
                    print("Vibranium: Full screen disabled")
                    speak("Full screen disabled")
                elif "open miniplayer mode".lower() in query.lower():
                    pyautogui.press("i")
                    print("Vibranium: Miniplayer mode enabled")
                    speak("Miniplayer mode enabled")
                elif "expand miniplayer mode".lower() in query.lower():
                    pyautogui.press("i")
                    print("Vibranium: Miniplayer mode expanded")
                    speak("Miniplayer mode expanded")
                elif "open caption".lower() in query.lower():
                    pyautogui.press("c")
                    print("Vibranium: Captions enabled")
                    speak("Captions enabled")
                elif "close caption".lower() in query.lower():
                    pyautogui.press("c")
                    print("Vibranium: Captions disabled")
                    speak("Captions disabled")
                elif "volume up".lower() in query.lower():
                    from keyboard import volumeup
                    print("Vibranium: Increasing volume sir")
                    speak("Increasing volume sir")
                    volumeup()
                elif "volume down".lower() in query.lower():
                    from keyboard import volumedown
                    print("Vibranium: Decreasing volume sir")
                    speak("Decreasing volume sir")
                    volumedown()

                elif "open camera".lower() in query.lower():
                    print("Vibranium: Opening camera sir")
                    speak("Opening camera sir")
                    cap = cv2.VideoCapture(0)
                    ret, img = cap.read()
                    if ret:
                      cv2.imshow('webcam', img)
                      filename = f"Image_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
                      cv2.imwrite(filename, img)
                      print(f"Vibranium: Okay sir, image saved as {filename}")
                      speak("Okay sir, image is saved on your device")
                      cv2.waitKey(10000)
                    else:
                      print("Vibranium: Sorry sir, could not access the camera")
                      speak("Sorry sir, I could not access the camera")
                    cap.release()
                    cv2.destroyAllWindows()

                elif "switch the window".lower() in query.lower():
                    pyautogui.keyDown("alt")
                    pyautogui.press("tab")
                    time.sleep(1)
                    pyautogui.keyUp("alt")

                elif "tell me the news".lower() in query.lower():
                    print("Vibranium: Please wait sir, collecting the news")
                    speak("Please wait sir, collecting the news")
                    news()

                elif "what is the weather today in".lower() in query.lower():
                    city = query.lower().split("in ")[-1].strip()
                    try:
                        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={OWM_KEY}&units=metric"
                        response = requests.get(url).json()
                        print(response)  # Debug: Print the raw API response
                        if str(response.get("cod")) == "200":
                            temp = response["main"]["temp"]
                            feels = response["main"]["feels_like"]
                            desc = response["weather"][0]["description"]
                            result = f"The weather in {city} is {desc}, {temp} degrees Celsius, feeling like {feels} degrees."
                            print(f"Vibranium: {result}")
                            speak(result)
                        else:
                            print(f"Vibranium: Sorry sir, I couldn't find weather data for {city}.")
                            speak(f"Sorry sir, I couldn't find weather data for {city}.")
                    except Exception as e:
                        print(f"Weather Error: {e}")
                        speak("Sorry sir, I was unable to get the weather.")

                elif "set an alarm".lower() in query.lower():
                    print("Vibranium: Please tell or type the time (e.g. 10 and 30 and 00)")
                    speak("Please tell or type the time sir")
                    a = get_alarm_input()
                    if a:
                       alarm(a)
                       print("Vibranium: Alarm is set, sir")
                       speak("Alarm is set, sir")
                    else:
                       print("Vibranium: Sorry sir, I couldn't get the time")
                       speak("Sorry sir, I couldn't get the time")

                elif "play music".lower() in query.lower():
                  music = r"F:\Software Back up\Entertainment\Songs\Music"
                  if os.path.exists(music):
                      songs = os.listdir(music)
                      mp3_songs = [s for s in songs if s.endswith('.mp3')]
                      if mp3_songs:
                          print("Vibranium: Playing music sir")
                          speak("Playing music sir")
                          os.startfile(os.path.join(music, mp3_songs[0]))
                      else:
                          print("Vibranium: No mp3 files found, sir.")
                          speak("No mp3 files found, sir.")
                  else:
                      print("Vibranium: Music directory not found, sir.")
                      speak("Music directory not found, sir.")

                elif "I am tired".lower() in query.lower():
                    print("Vibranium: Playing your favourite songs sir")
                    speak("Playing your favourite songs")
                    webbrowser.open("https://open.spotify.com/playlist/2SNI12ibWSwJ4gJazGPz0X")

                elif "tell me a joke".lower() in query.lower():
                    joke = pyjokes.get_joke()
                    print(f"Vibranium: {joke}")
                    speak(joke)

                elif "take a screenshot".lower() in query.lower():
                    print("Vibranium: Okay sir, taking a screenshot")
                    speak("Okay sir, taking a screenshot")
                    time.sleep(5)
                    filename = f"Screenshot_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
                    image = pyautogui.screenshot()
                    image.save(filename)
                    print(f"Vibranium: Screenshot saved as {filename}")
                    speak(f"Screenshot saved sir")

                elif "shutdown the system".lower() in query.lower():
                    print("Shutting down the system")
                    speak("Shutting down the system")
                    os.system("shutdown /s /t 5")

                elif "restart the system".lower() in query.lower():
                    print("Restarting the system")
                    speak("Restarting the system")
                    os.system("shutdown /r /t 5")

                elif "sleep the system".lower() in query.lower():
                    print("Putting the system to sleep")
                    speak("Putting the system to sleep")
                    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

                elif "quit".lower() in query.lower() or "exit".lower() in query.lower() or "Vibranium quit".lower() in query.lower():
                    print("Ok sir. Quitting. Have a nice day sir.")
                    speak("Ok sir. Quitting. Have a nice day sir.")
                    sys.exit()
                
                elif "open chat mode".lower() in query.lower() or "text mode".lower() in query.lower():
                    speak("Entering chat mode, sir. The microphone is paused.")
                    print("\n" + "="*50)
                    print("🔴 VIBRANIUM CHAT MODE ACTIVATED")
                    print("👉 Paste your long questions or text below.")
                    print("👉 Press Ctrl+Z (then Enter) on a new line to submit your query.")
                    print("👉 Type 'exit' and press Ctrl+Z to switch back to voice control.")
                    print("="*50 + "\n")
                    
                    while True:
                        print("You (Paste text + Ctrl+Z + Enter):")
                        lines = []
                        while True:
                            try:
                                line = input()
                            except EOFError:
                                break
                            lines.append(line)
                        
                        text_query = "\n".join(lines).strip()
                        
                        # Check if the user wants to leave chat mode
                        if text_query.lower() in ["exit", "voice mode", "close chat mode", "quit chat"]:
                            print("\n❌ Exiting Chat Mode...")
                            speak("Exiting chat mode and activating voice control, sir.")
                            print("🎙️ VOICE CONTROL READY...\n")
                            break
                            
                        # Ignore empty submissions
                        if not text_query:
                            continue
                            
                        print("\nThinking...")
                        # Send the long query to your existing Gemini chat function
                        chat(text_query)
                        print("-" * 50)

                else:
                    sites = load_sites()
                    matched_site = False
                    for site in sites:
                        if f"Open {site[0]}".lower() in query.lower():
                            print(f"Vibranium: Opening {site[0]} sir")
                            speak(f"Opening {site[0]} sir")
                            webbrowser.open(site[1])
                            matched_site = True
                            break  # Exits the site loop cleanly
                        elif f"Close {site[0]}".lower() in query.lower():
                            print(f"Vibranium: Closing {site[0]} sir")
                            speak(f"Closing {site[0]} sir")
                            pyautogui.hotkey('ctrl', 'w')
                            matched_site = True
                            break  # Exits the site loop cleanly

                    # 3. Gemini Fallback (Only runs if NO other command or website matched)
                    if not matched_site:
                        chat(query)