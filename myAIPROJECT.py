# INSTALL ALL THE PAKAGES
import os
import subprocess
import speech_recognition as sr
import webbrowser
import pyttsx3
import win32com.client
import datetime
import cohere
speaker = win32com.client.Dispatch("SAPI.spvoice")
chatstr = ""
# ARTIFICIAL INTELLIGENCE
def ai(promp):
    co = cohere.Client(
        api_key="ENTER YOUR API",  # This is your trial API key
    )
    stream = co.chat_stream(
        model='command-r-plus',
        message=promp,
        temperature=0.3,
        chat_history=[{"role": "User",
                       "message": promp}],
        prompt_truncation='AUTO',
        connectors=[{"id": "web-search"}]
    )

    for event in stream:
        if event.event_type == "text-generation":
            print(event.text,end='')

# proessing user command
def say(query):
    engine = pyttsx3.init()
    engine.say(query)
    engine.runAndWait()

# taking command from user
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("recognizing....")
            query = r.recognize_google(audio, language="en-in")
            print(f"user said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from jarvis"

# main function
if __name__ == '__main__':
    print('pychram')
    say("hello i am jarvis")
    while True:
        print("listening....")
        query = takeCommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"], ["w3schools", "https://www.w3schools.com/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening{site[0]} sir...")
                webbrowser.open(site[1])

        # say(query)
        if "the time" in query:
            now = datetime.datetime.now()

            # Print the current date and time
            print("Current date and time:", now)

            # Format the date and time
            formatted_date = now.strftime("%Y-%m-%d %H:%M:%S")
            say(f"sir the time is {formatted_date}")
        if "open video ".lower() in query.lower():
            os.startfile(r"C:\Users\jeetd\Videos\Captures\K.G.F CHAPTER 2.mkv")
        if "play music".lower() in query.lower():
            os.startfile(r"C:\Users\jeetd\Music\8 Parche [Lofi Song] Baani Sandhu _ Slowed   Reverb _ 8D Audio _ Bollywood Lofi Song _ Punjabi Songs(MP3_70K).mp3")
        if "open word".lower() in query.lower():
            print("Attempting to open Word...")
            subprocess.run([r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"])
            print("Command executed.")
        if "using artificial intelligence".lower() in query.lower():
            ai(promp=query)
