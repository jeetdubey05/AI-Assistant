# INSTALL THESE PAKAGES
import datetime
import os
import subprocess
import webbrowser
import cohere
import pyttsx3
import win32com.client
import cv2
import numpy as np
import pywhatkit as kit
import pyautogui
import time
import speech_recognition as sr
import json
import requests
import time
from newsapi import NewsApiClient
import pyautogui
from plyer import notification
import screen_brightness_control as sbc
from datetime import date, timedelta
import re

speaker = win32com.client.Dispatch("SAPI.spvoice")
chatStr = ""


def fetch_weather(location):
    api_key = 'your API KEy'
    try:
        # Fetch current weather data
        current_url = f'http://api.weatherapi.com/v1/current.json?key={api_key}&q={location}'
        current_response = requests.get(current_url)
        current_response.raise_for_status()  # Raises error for bad status

        current_weather_data = current_response.json()
        current = current_weather_data['current']
        current_temp = current['temp_c']
        current_condition = current['condition']['text']
        current_humidity = current['humidity']
        current_wind_speed = current['wind_kph']
        current_wind_direction = current['wind_dir']
        current_feels_like = current['feelslike_c']
        current_uv_index = current['uv']
        current_precipitation = current['precip_mm']

        print(f"Location: {location}")
        print(f"Current Temperature: {current_temp}°C")
        print(f"Current Condition: {current_condition}")
        print(f"Current Humidity: {current_humidity}%")
        print(f"Current Wind Speed: {current_wind_speed} kph")
        print(f"Current Wind Direction: {current_wind_direction}")
        print(f"Current Feels Like: {current_feels_like}°C")
        print(f"Current UV Index: {current_uv_index}")
        print(f"Current Precipitation: {current_precipitation} mm")

        # Fetch tomorrow's weather data
        forecast_url = f'http://api.weatherapi.com/v1/forecast.json?key={api_key}&q={location}&days=2'
        forecast_response = requests.get(forecast_url)
        forecast_response.raise_for_status()  # Raises error for bad status

        forecast_weather_data = forecast_response.json()
        forecast = forecast_weather_data['forecast']['forecastday'][1]
        forecast_temp = forecast['day']['avgtemp_c']
        forecast_condition = forecast['day']['condition']['text']
        forecast_humidity = forecast['day']['avghumidity']
        forecast_wind_speed = forecast['day']['maxwind_kph']
        forecast_precipitation = forecast['day']['totalprecip_mm']

        print("\nTomorrow's Weather:")
        print(f"Location: {location}")
        print(f"Average Temperature: {forecast_temp}°C")
        print(f"Condition: {forecast_condition}")
        print(f"Humidity: {forecast_humidity}%")
        print(f"Wind Speed: {forecast_wind_speed} kph")
        print(f"Precipitation: {forecast_precipitation} mm")

    except requests.exceptions.RequestException as e:
        print(f"Error fetching weather data: {e}")
        # Retry after 5 seconds if an error occurs
        time.sleep(5)


def chat(promp):
    text = ""
    global chatStr, event
    # print(chatStr,end='')
    # say(chatStr)
    co = cohere.Client(
        api_key="your API key",  # This is your trial API key
    )
    chatStr += f"user:{promp}\n jarvis:"
    stream = co.chat_stream(
        model='command-r-plus',
        message=chatStr,
        temperature=0.3,

        prompt_truncation='AUTO',
        connectors=[{"id": "web-search"}]
    )
    try:
        print("processing....", flush=True)
        final_text = ""
        for event in stream:
            if event.event_type == "text-generation":
                chatStr += f"{event.text}\n"
                final_text += event.text

        print(final_text)
        say(final_text)
        return event.text
    except Exception as e:
        return "Some Error Occurred. Sorry from jarvis"


# ARTIFICIAL INTELLIGENCE
text = ""

def screen():
    # take screenshort
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    folder_path=r'your folderpath for saving screenshoth'
    image_name=f"screenshot_{current_time}.png"

    screenshot = pyautogui.screenshot()
    save_path = os.path.join(folder_path, image_name)
    screenshot.save(save_path)
    
    # Sending a desktop notification
    notification.notify(
        title="Screenshot Taken",
        message=f"Your screenshot has been saved as screenshot_{current_time}.png",
        timeout=5  # notification timeout in seconds
    )


def ai(promp):
    text = ""
    co = cohere.Client(
        api_key="your API key",  # This is your trial API key
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
    try:
        print("GENERATING RESULT....")
        say("GENERATING RESULT")
        for event in stream:
            if event.event_type == "text-generation":
                text += event.text

        # print(text, end='')

        def split_into_sentences(text):
            """
            Splits the input text into sentences based on full stops.
            """
            sentences = text.split('. ')
            return sentences

        # Example usage:
        user_text = text
        sentences_list = split_into_sentences(user_text)

        for sentence in sentences_list:
            print(sentence)
            say(sentence)
    except Exception as e:
        return "Some Error Occurred. Sorry from jarvis"


# processing user command
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
            return "Some Error Occurred. Sorry"


def msg():
    # File to store contacts
    contacts_file = 'contacts.json'

    # Load contacts from file
    def load_contacts():
        try:
            with open(contacts_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    # Save contacts to file
    def save_contacts():
        with open(contacts_file, 'w') as f:
            json.dump(contacts, f)

    # Dictionary to store contacts
    contacts = {k.lower(): v for k, v in load_contacts().items()}

    def listen_for_contact_name():
        recognizer = sr.Recognizer()
        while True:
            with sr.Microphone() as source:
                print("Listening for the contact name or say 'new contact' to add a new one or say exit to exit from messaging ...")
                say("Listening for the contact name or say 'new contact' to add a new one or say exit to exit from messaging ...")
                audio = recognizer.listen(source)
            try:
                contact_name = recognizer.recognize_google(audio).lower()
                if contact_name == "exit":
                    return "exit"
                print(f"Recognized contact name: {contact_name}")
                return contact_name
            except sr.UnknownValueError:
                print("Could not understand the contact name. Please try again.")
            except sr.RequestError as e:
                print(f"Could not request results; {e}")
                break  # Exit the loop if there's a request error

    def listen_and_confirm():
        recognizer = sr.Recognizer()
        while True:
            with sr.Microphone() as source:
                print("Listening for your message...")
                audio = recognizer.listen(source)
            try:
                message = recognizer.recognize_google(audio)
                print(f"Recognized message: {message}")
                # Voice confirmation
                print("Please say 'send' to confirm or 'back"
                      "' to re-record the message.")
                while True:
                    with sr.Microphone() as source:
                        audio = recognizer.listen(source)
                        confirmation = recognizer.recognize_google(audio).lower()
                        if "send" in confirmation:
                            return message
                        elif "back" in confirmation:
                            print("Please re-record your message.")
                            break
                        else:
                            print("Could not understand confirmation. Please try again.")
            except sr.UnknownValueError:
                print("Could not understand audio. Please try again.")
            except sr.RequestError as e:
                print(f"Could not request results; {e}")
                break  # Exit the loop if there's a request error

    def add_new_contact():

        recognizer = sr.Recognizer()
        while True:
            with sr.Microphone() as source:
                print("Please say the new contact's name...")
                audio = recognizer.listen(source)
                try:
                    name = recognizer.recognize_google(audio).lower()
                    print(f"Recognized name: {name}")
                    break
                except sr.UnknownValueError:
                    print("Could not understand the name. Please try again.")
                except sr.RequestError as e:
                    print(f"Could not request results; {e}")
                    break

        while True:
            with sr.Microphone() as source:
                print("Please say the new contact's phone number (with country code)...")
                audio = recognizer.listen(source)
                try:
                    number = recognizer.recognize_google(audio).replace(" ", "")
                    print(f"Recognized number: {number}")
                    break
                except sr.UnknownValueError:
                    print("Could not understand the phone number. Please try again.")
                except sr.RequestError as e:
                    print(f"Could not request results; {e}")
                    break

        contacts[name] = number
        save_contacts()  # Save contacts to file
        print(f"Contact '{name}' added successfully.")

    def send_whatsapp_message(contact_name, message):
        if contact_name == "new contact":
            add_new_contact()
            return

        if contact_name not in contacts:
            print(f"Contact '{contact_name}' not found.")
            return

        phone_number = contacts[contact_name]
        kit.sendwhatmsg_instantly(phone_number, message, wait_time=10, tab_close=False)
        time.sleep(15)
        pyautogui.press('enter')
        print(f"Message sent to {contact_name}")
        say(f"Message sent to {contact_name}")


    # Main loop
    while True:
        contact_name = listen_for_contact_name()

        if contact_name == "exit":
            print("Exiting...")
            break

        if contact_name == "new contact":
            add_new_contact()
        else:
            message = listen_and_confirm()
            if message:
                send_whatsapp_message(contact_name, message)


def obj():
    net = cv2.dnn.readNet("yolov4-tiny.weights", "yolov4-tiny.cfg")
    layer_names = net.getLayerNames()
    output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers()]

    # Load the COCO labels
    with open("coco.names", "r") as f:
        classes = [line.strip() for line in f.readlines()]

    # Assign a color to each class
    colors = np.random.uniform(0, 255, size=(len(classes), 3))

    # Start video capture
    cap = cv2.VideoCapture(0)
    detected_labels_dict = {}
    say("Welcome to object detection")

    while True:
        ret, frame = cap.read()
        height, width, channels = frame.shape

        # Prepare the frame for YOLO
        blob = cv2.dnn.blobFromImage(frame, 0.00392, (416, 416), (0, 0, 0), True, crop=False)
        net.setInput(blob)
        outs = net.forward(output_layers)

        # Process the results
        class_ids = []
        confidences = []
        boxes = []

        for out in outs:
            for detection in out:
                scores = detection[5:]
                class_id = np.argmax(scores)
                confidence = scores[class_id]
                if confidence > 0.5:
                    center_x = int(detection[0] * width)
                    center_y = int(detection[1] * height)
                    w = int(detection[2] * width)
                    h = int(detection[3] * height)
                    x = int(center_x - w / 2)
                    y = int(center_y - h / 2)
                    boxes.append([x, y, w, h])
                    confidences.append(float(confidence))
                    class_ids.append(class_id)

        indexes = cv2.dnn.NMSBoxes(boxes, confidences, 0.5, 0.4)

        if len(indexes) > 0:
            for i in indexes.flatten():
                x, y, w, h = boxes[i]
                label = str(classes[class_ids[i]])
                confidence = confidences[i] * 100
                label_with_confidence = f"{label} {confidence:.2f}%"
                color = colors[class_ids[i]].tolist()
                if label not in detected_labels_dict or detected_labels_dict[label] < confidence:
                    detected_labels_dict[label] = confidence
                cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)
                cv2.putText(frame, label_with_confidence, (x, y + 30), cv2.FONT_HERSHEY_PLAIN, 2, color, 2)

        # Display the frame
        cv2.imshow('YOLOv4-Tiny Live Detection', frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

    # Display and say detected labels with the highest confidence levels
    detected_labels_list = [f"{label} {conf:.2f}%" for label, conf in detected_labels_dict.items()]

    print("Detected objects with confidence levels:", detected_labels_list)
    say("Detected objects with confidence levels: " + ", ".join(detected_labels_list))
    promp = ("Describe the objects : " + ", ".join(detected_labels_list))
    ai(promp)

def set_brightness(input_str):
    # Regular expression to match phrases like "set brightness to", "make brightness to", etc.
    match = re.search(r'brightness (?:to|at|set to|make to|adjust to) (\d+)', input_str, re.IGNORECASE)
    if match:
        brightness_level = int(match.group(1))
        if 0 <= brightness_level <= 100:
            # Set the screen brightness
            sbc.set_brightness(brightness_level)
            print(f"Screen brightness set to {brightness_level}%")
        else:
            print("Please provide a brightness level between 0 and 100.")
    else:
        print("Invalid input format. Please use phrases like 'set brightness to X percent'.")




def news():
    # Initialize NewsApiClient
    newsapi = NewsApiClient(api_key='your API key')
    print("NewsAPI client initialized")

    # This can be replaced with a function call or input method as needed
    say("please tell me the catregory"
        "")
    category = takeCommand()
    # Expand the list of Indian news sources and domains
    indian_sources = 'the-hindu,ndtv,aajtak,india-today,hindustan-times,zee-news,times-of-india,republic-tv,news18,firstpost,economic-times,financial-express,indian-express'
    indian_domains = 'thehindu.com,ndtv.com,aajtak.in,indiatoday.in,hindustantimes.com,zeenews.india.com,timesofindia.indiatimes.com,republicworld.com,news18.com,firstpost.com,economictimes.indiatimes.com,financialexpress.com,indianexpress.com'

    def fetch_articles(query, sources=None, domains=None, page_size=5, from_date=None, to_date=None):
        if sources:
            return newsapi.get_top_headlines(
                q=query,
                sources=sources,
                language='en',
                page_size=page_size
            )
        else:
            return newsapi.get_everything(
                q=query,
                domains=domains,
                from_param=from_date,
                to=to_date,
                language='en',
                sort_by='relevancy',
                page_size=page_size
            )

    def print_articles(articles, header):
        print(header)
        for article in articles:
            title = article.get('title', 'No Title')
            description = article.get('description', 'No Description')
            url = article.get('url', 'No URL')
            print(f"Title: {title}\nDescription: {description}\nURL: {url}\n")

    # Fetch top headlines from specific Indian sources
    top_headlines = fetch_articles(query=category, sources=indian_sources)
    print("Fetched top headlines")

    # Fetch all articles with an adjusted date range and specific domains
    today = date.today()  # Correctly get today's date
    start_date = (today - timedelta(days=30)).strftime('%Y-%m-%d')
    end_date = today.strftime('%Y-%m-%d')

    all_articles = fetch_articles(query=category, domains=indian_domains, from_date=start_date, to_date=end_date)
    print("Fetched all articles")

    # Print results
    print_articles(top_headlines['articles'], "Top Headlines:\n")
    print_articles(all_articles['articles'], "All Articles:\n")




if __name__ == '__main__':
    print('Pycharm')
    say("Hello, I am Jarvis")

    while True:
        print("Listening....")
        query = takeCommand()

        sites = [
            ["youtube", "https://www.youtube.com"],
            ["wikipedia", "https://www.wikipedia.com"],
            ["google", "https://www.google.com"],
            ["w3schools", "https://www.w3schools.com/"]
        ]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]}, sir...")
                webbrowser.open(site[1])
                break

        if "the time" in query:
            now = datetime.datetime.now()

            # Print the current date and time
            print("Current date and time:", now)

            # Format the date and time
            formatted_date = now.strftime("%Y-%m-%d %H:%M:%S")
            say(f"sir the time is {formatted_date}")
        elif "open video" in query:
            os.startfile(r"C:\Users\jeetd\Videos\dolly.mp4")

        elif "play music" in query:
            os.startfile(
                r"C:\Users\jeetd\Music\8 Parche [Lofi Song] Baani Sandhu _ Slowed   Reverb _ 8D Audio _ Bollywood Lofi Song _ Punjabi Songs(MP3_70K).mp3")

        elif "open".lower() in query.lower():
            apps = [
                ["Excel", r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"],
                ["powerpoint", r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"],
                ["WhatsApp", r"explorer.exe shell:AppsFolder\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"],
                ["edge", r"explorer.exe shell:AppsFolder\MSEdge"],
                ["Word", r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"],
            ]
            for app in apps:
                if f"open {app[0]}".lower() in query.lower():
                    say(f"Opening {app[0]}, sir...")
                    subprocess.run(app[1])
                    print("Command executed.")
                    break

        elif "using" in query:
            query = query.replace("using", "").strip()
            ai(promp=query)

        elif "object scanning".lower() in query.lower():
            obj()

        elif "send message".lower() in query.lower():
            msg()

        elif "news" in query.lower():
            news()
        elif "screenshot" in query:
            screen()
        elif "brightness" in query:
            user_input = query
            set_brightness(user_input)

        elif "weather" in query:
            try:
                say("Please tell the location")
                location = takeCommand()
                fetch_weather(location)
                say("Data has been fetched")
                print(f"Recognized location: {location}")
            except Exception as e:
                print(f"Error: {e}")

        else:
            chat(promp=query)
