import threading
import openai
import speech_recognition as sr
import os
import webbrowser
import datetime
import win32com.client

chatStr = ""
stop_jarvis = False

def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = "add-your-api-keys" 
    chatStr += f"Sir: {query}\n Jarvis: "
    response = openai.Completion.create(
        model="gpt-3.5-turbo-instruct",
        prompt=chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    speaker.Speak(response.choices[0].text)
    chatStr += f"{response.choices[0].text}\n"
    return response.choices[0].text

def ai(prompt):
    openai.api_key = "add-your-api-keys" 
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"
    response = openai.Completion.create(
        model="gpt-3.5-turbo-instruct",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    text += response.choices[0].text
    if not os.path.exists("Openai"):
        os.mkdir("Openai")
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)

# Selecting a female voice
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices("Name=Microsoft Hazel Desktop").Item(0)  # Change "Microsoft Hazel Desktop" to the desired female voice

def takeCommand():
    global stop_jarvis
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            if "stop" in query.lower():
                stop_jarvis = True
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

def speak_response(response):
    global stop_jarvis
    for sentence in response.split("."):
        if stop_jarvis:
            break
        speaker.Speak(sentence.strip())
    stop_jarvis = False

if __name__ == '__main__':
    print('Welcome to Jarvis AI')
    speaker.Speak("Jarvis AI")
    while True:
        print("Listening...")
        query = takeCommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"], ]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
                exit()
        if "music" in query:
            musicPath = "https://open.spotify.com/track/0F7FA14euOIX8KcbEturGH?si=5faa8772d7f940d0"
            os.system(f"start {musicPath}")
            exit()
        elif "the time" in query:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            speaker.Speak(f"Sir time is {hour} HOUR {min} minutes")
        elif "firefox".lower() in query.lower():
            os.system(r'"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"') 
        elif "date" in query:
            current_date = datetime.datetime.now().strftime("%A, %B %d, %Y")
            speaker.Speak(f"Sir, today is {current_date}")

        elif "reset chat".lower() in query.lower():
            chatStr = ""
        elif "stop".lower() in query.lower():
            exit()
        else:
            print("Chatting...")
            response = chat(query)
