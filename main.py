# You can use webdriver package for extra functionality
import speech_recognition as sr
import win32com.client
import webbrowser
import numpy as np

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    print(text)
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='en-in')
            return(str(query))
        except Exception as e:
            return "Sorry sir some error occured"
            
def openApps(app):
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.Run(f"{app}")

if __name__  == "__main__":
    say("Hello sir, I am Jarvis, How can I help you?")
    while 1:
        query = takeCommand()

        sites = np.array([
            ["youtube", "https://youtube.com"],
            ["wikipedia", "https://wikipedia.com"],
            ["google", "https://google.com"]
        ])

        applications = np.array([
            ["spotify", "spotify"],
            ["word", "winword"],
            ["powerpoint", "POWERPNT"],
            ["chrome", "chrome"],
            ["postman", "postman"],
            ["code", "code"],
        ])

        for site in sites:
            if f"Jarvis open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir..")
                webbrowser.open(site[1])
                break
        
        for application in applications:
            if f"Jarvis run {application[0]}".lower() in query.lower():
                say(f'Running {application[0]}...')
                openApps(application[1])
                break

        if "Jarvis stop server".lower() in query.lower():
            say("Stopping the server sir..")
            exit()