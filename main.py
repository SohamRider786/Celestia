import datetime
import webbrowser
import speech_recognition as sr
import wikipedia
import win32com.client
import winapps
from AppOpener import open as op
import os
import openai
from env_var import API_KEY
from env_var import APP_ID
from env_var import AUTH_TOKEN
from env_var import ACCOUNT_SID
from env_var import X_RAPIDAPI_KEY
from env_var import NEWS_API
import requests
import geocoder
from datetime import date
from deep_translator import GoogleTranslator
import smtplib
from twilio.rest import Client
import wolframalpha

speaker = win32com.client.Dispatch("SAPI.SpVoice")
client = wolframalpha.Client(APP_ID)


def calculate(text1):
    ind = text1.lower().split().index('calculate')
    text1 = text1.split()[ind + 1:]
    response = client.query(" ".join(text1))
    answer = next(response.results).text
    print("Celestia:the answer is= ", answer)
    speaker.Speak(f"the answer is={answer}")


def send_message():
    print("Celestia:Enter the receiver number:")
    speaker.Speak("Enter the receiver number:")
    number = input()
    print("Enter the message:")
    speaker.Speak("Enter the message:")
    message = input()
    cl = Client(ACCOUNT_SID, AUTH_TOKEN)
    message = cl.messages.create(
        body=message,
        from_='+13343779640',
        to=number
    )
    print(message.sid)
    print("Message sent successfully")
    speaker.Speak("Message sent successfully")


def sent_email():
    print("Receiver Email:")
    msg = ''
    speaker.Speak("Receiver Email:")
    to = input()
    print("Subject:")
    speaker.Speak("Subject(speak up!):")
    sub = takeCommand()
    send_mode = int(input("enter mode:\n 0->AI generated mail\n 1->manual mail\n"))
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('chowdhurysoham26@gmail.com', 'qbha eouv vmup ekpf')
    if send_mode == 0:
        print("topic:")
        speaker.Speak("topic:")
        topic = takeCommand()
        result = ai(prompt="write an email on" + topic)
        msg = f"Subject:{sub}\n\n{result}"
    elif send_mode == 1:
        print("Enter the message:")
        speaker.Speak("Enter the message:")
        msg = "Subject:" + sub + "\n\n" + takeCommand()
    server.sendmail('chowdhurysoham26@gmail.com', to, msg)
    server.quit()
    print("Email sent successfully")
    speaker.Speak("Email sent successfully")


def search_google(query):
    url = "https://www.google.com/search?q=" + query
    webbrowser.open(url)


def search_youtube(speech):
    if "youtube" in speech:
        search = speech.replace("youtube", "")
        print("Celestia:Searching YouTube for:", search)
        speaker.Speak(f"Searching YouTube for:{search}")
        webbrowser.open(f"https://www.youtube.com/results?search_query={search}")


def get_translate(speech):
    print("Celestia:Enter the language you want to translate to")
    speaker.Speak("Enter the language you want to translate to")
    dest_language = takeCommand()
    languages = {
        'arabic': 'ar',
        'bengali': 'bn',
        'chinese (simplified)': 'zh-cn',
        'chinese (traditional)': 'zh-tw',
        'croatian': 'hr',
        'dutch': 'nl',
        'english': 'en',
        'filipino': 'tl',
        'french': 'fr',
        'german': 'de',
        'greek': 'el',
        'gujarati': 'gu',
        'hindi': 'hi',
        'italian': 'it',
        'japanese': 'ja',
        'korean': 'ko',
        'latin': 'la',
        'malayalam': 'ml',
        'marathi': 'mr',
        'mongolian': 'mn',
        'myanmar (burmese)': 'my',
        'nepali': 'ne',
        'norwegian': 'no',
        'odia': 'or',
        'portuguese': 'pt',
        'punjabi': 'pa',
        'russian': 'ru',
        'spanish': 'es',
        'tamil': 'ta',
        'telugu': 'te'}
    if dest_language in languages.keys():
        destlang = languages.get(dest_language)
    else:
        print("Celestia:Sorry, I don't know that language")
        speaker.Speak("Sorry, I don't know that language")
        return
    print("Celestia:language code=", destlang)
    to_translate = speech
    translated = GoogleTranslator(source='auto', target=destlang).translate(to_translate)
    print("Celestia:The translation is", translated)
    speaker.Speak(translated)


def play_music():
    song_dir = "F:\Music"
    songs = os.listdir(song_dir)
    print("Celestia:Type the name of the song that you want to be played")
    speaker.Speak(f"Type the name of the song that you want to be played")
    song = input("user said:")
    try:
        for it in songs:
            if song in it.lower():
                print("Celestia:Playing song:", it)
                speaker.Speak(f"Playing song:{it}")
                os.startfile(os.path.join(song_dir, it))
    except Exception as e:
        print("Song not found or some error occured:", e)


def get_date():
    Date = date.today().strftime("%d/%m/%Y")
    print(f"Celestia:The date is {Date}")
    speaker.Speak(f"The date is {Date}")


def get_meaning(text3):
    if "what is the meaning of" in text3:
        text3 = text3.replace("what is the meaning of", "")
    elif "what do you mean by" in text3:
        text3 = text3.replace("what do you mean by", "")
    result = wikipedia.summary(text3, sentences=3)
    print(f"Celestia:{result}")
    speaker.Speak(result)


def get_location():
    location = geocoder.ip("me")
    print("The latitude of the location is: ", location.latlng[0])
    print("The longitude of the location is: ", location.latlng[1])
    print(location.address)


def get_weather():
    location = geocoder.ip("me")
    url = "https://weatherapi-com.p.rapidapi.com/current.json"

    querystring = {"q": f"{location.latlng[0]},{location.latlng[1]}"}

    headers = {
        "X-RapidAPI-Key": X_RAPIDAPI_KEY,
        "X-RapidAPI-Host": "weatherapi-com.p.rapidapi.com"
    }

    response = requests.get(url, headers=headers, params=querystring)
    print("Celestia:")
    print("Temperature in celsius is: ", response.json()["current"]["temp_c"])
    speaker.Speak(f"temperature in celsius is: {response.json()['current']['temp_c']}")
    print("condition is: ", response.json()["current"]["condition"]["text"])
    speaker.Speak(f"condition is: {response.json()['current']['condition']['text']}")
    print("feels like is: ", response.json()["current"]["feelslike_c"])
    speaker.Speak(f"feels like is: {response.json()['current']['feelslike_c']}")
    print("wind direction is: ", response.json()["current"]["wind_dir"])
    speaker.Speak(f"wind direction is: {response.json()['current']['wind_dir']}")
    print("humidity is: ", response.json()["current"]["humidity"])
    speaker.Speak(f"humidity is: {response.json()['current']['humidity']}")


def get_jokes():
    url = "https://daddyjokes.p.rapidapi.com/random"

    headers = {
        "X-RapidAPI-Key": X_RAPIDAPI_KEY,
        "X-RapidAPI-Host": "daddyjokes.p.rapidapi.com"
    }

    response = requests.get(url, headers=headers)
    print("Celestia:", response.json()['joke'])
    speaker.Speak(response.json()['joke'])


def get_news():
    response = requests.get(NEWS_API)
    # for each "articles" in response.json():
    #     print(each["title"])
    #     print(each["description"])
    for each in response.json()["articles"]:
        print(f"Celestia:{each['title']}")
        speaker.Speak(each["title"])
        print(f"Celestia:{each['description']}")
        speaker.Speak(each["description"])
        print("\n***********************************************************\n")


def ai(prompt):
    openai.api_key = API_KEY
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "system",
                "content": prompt
            },
            {
                "role": "user",
                "content": ""
            }
        ],
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        result = ""
        result += response["choices"][0]["message"]["content"]
        print("Celestia:", result)
        speaker.Speak(f"{result}")
        if not os.path.exists("Openai"):
            os.mkdir("Openai")
        file = open("Openai/prompt.txt", "a+")
        file.write(f"user:{prompt}")
        file.write(f"Celestia:{result}")
        return result
    except Exception as e:
        print("Celestia:Error in AI", e)
        speaker.Speak("Error in AI")


def open_websites(text3):
    site = ""
    words = text3.split()
    pos = 0
    for word in words:
        if word.lower() == "Open".lower():
            try:
                site = words[pos + 1]
            except Exception as e:
                print(f"Invalid command", e)
                speaker.Speak("Invalid Instruction")
        pos = pos + 1
        site = site
    speaker.speak(f"Opening {site}")
    webbrowser.open(site + ".com")


def get_time():
    strfTime = datetime.datetime.now().strftime("%H:%M:%S")
    print("Celestia:The time is", strfTime)
    speaker.Speak(f"The time is {strfTime}")


def open_application(text2):
    for apks in winapps.list_installed():
        if f"Open {apks.name}".lower() in text2.lower():
            print(f"Opening ", apks.name)
            speaker.Speak(f"Opening {apks.name}")
            op(apks.name)
            return
    open_websites(text2)


def do_tasks(text1):
    if "open" in text1:
        open_application(text1)
        # open_websites(text)
    elif "what is the date today" in text1:
        get_date()
    elif "what is the time" in text1:
        get_time()
    elif "what is the news" in text1:
        get_news()
    elif "what is the weather" in text1:
        get_weather()
    elif "what is my location" in text1:
        get_location()
    elif "tell me a joke" in text1:
        get_jokes()
    elif "calculate" in text1:
        calculate(text1)
    elif "what is the meaning of" in text1 or "what do you mean by" in text1:
        get_meaning(text1)
    elif "play music" in text1 or "play song" in text1:
        play_music()
    elif "translate" in text1:
        text1 = text1.replace("translate", "")
        get_translate(text1)
    elif "youtube" in text1:
        search_youtube(text1)
    elif "search" in text1:
        search_google(text1)
    elif "send an email" in text1:
        sent_email()
    elif "send a message" in text1 or "send message" in text1:
        send_message()
    else:
        ai(prompt=text1)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognising. . . ")
            query = r.recognize_google(audio, language="en-in")
            print(f"user said:{query}")
            query = query.lower()
            return query
        except Exception as e:
            print(f"Celestia:Speech could not be recognised {e}")
            speaker.Speak("Speech could not be recognised")
            return "Speech could not be recognised"


if __name__ == '__main__':
    print("Hi i am Celestia")
    speaker.Speak("Hi i am Celestia")
    print("Celestia:What do you want me to do?")
    speaker.Speak("What do you want me to do?")
    while 1:
        print("listening . . .")
        text = takeCommand()
        if "Speech could not be recognised".lower() in text.lower():
            continue
        elif text.lower() == 'stop'.lower():
            print("Celestia:Shutting down . . .")
            speaker.Speak("Shutting down . . .")
            break
        else:
            do_tasks(text)
        # todo:Natural language processing
        # todo:selenium
        # todo: add gui
    print("Celestia:Shutdown Completed")
    speaker.Speak("Shutdown Competed")