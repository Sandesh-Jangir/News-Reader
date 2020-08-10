from win32com.client import Dispatch
import requests
import json
import time
def speak(str):
    speak = Dispatch("Sapi.spVoice")
    speak.Speak(str)


if __name__ == "__main__":
    i = 1
    print ("News for Today\n")
    speak("News for today")
    url = "https://newsapi.org/v2/top-headlines?sources=bbc-news&apiKey=4dc1ba9aacb74817981a63edf86556ad"
    news = requests.get(url).text
    news_py = json.loads(news)
    arts = news_py['articles']
    for article in arts:
        try:
            print(i, article["title"])
            speak(article["title"])
            i += 1
        except:
            speak("Unable to search news plese check your internet")
    print ("\n Thank You For Listining")
    speak("thank you for lisning")
    time.sleep(2)

