import random
import time
import pandas as pd
import speech_recognition as sr
import numpy as np
from textblob import TextBlob
import warnings
warnings.filterwarnings("ignore")

def data():
    scores=pd.read_excel("live_data.xlsx")
    scores=scores.set_index('Project Number')
    return scores

def voice_input_and_update(num,scores):
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()    
    print('How do you feel about the project!'.format(1))
    guess = recognize_speech_from_mic(recognizer, microphone)
    # if there was an error, stop the game
    if guess["error"]:
        print("ERROR: {}".format(guess["error"]))
    # show the user the transcription
    print("You said: {}".format(guess["transcription"]))
    
    #Update
    inp=guess["transcription"]
    inp = TextBlob(inp)
    inp = inp.correct()
    scores["Live Score"][num]*=(scores["Count"][num])
    scores["Count"][num]+=1
    scores["Live Score"][num]+=(inp.sentiment.polarity*5)
    scores["Live Score"][num]/=scores["Count"][num]
    scores.to_excel("live_data.xlsx")       
    scores.to_html('live_scores.html')

def recognize_speech_from_mic(recognizer, microphone):
    """Transcribe speech from recorded from `microphone`.

    Returns a dictionary with three keys:
    "success": a boolean indicating whether or not the API request was
               successful
    "error":   `None` if no error occured, otherwise a string containing
               an error message if the API could not be reached or
               speech was unrecognizable
    "transcription": `None` if speech could not be transcribed,
               otherwise a string containing the transcribed text
    """
    # check that recognizer and microphone arguments are appropriate type
    if not isinstance(recognizer, sr.Recognizer):
        raise TypeError("`recognizer` must be `Recognizer` instance")

    if not isinstance(microphone, sr.Microphone):
        raise TypeError("`microphone` must be `Microphone` instance")

    # adjust the recognizer sensitivity to ambient noise and record audio
    # from the microphone
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    # set up the response object
    response = {
        "success": True,
        "error": None,
        "transcription": None
    }

    # try recognizing the speech in the recording
    # if a RequestError or UnknownValueError exception is caught,
    #     update the response object accordingly
    try:
        response["transcription"] = recognizer.recognize_google(audio)
    except sr.RequestError:
        # API was unreachable or unresponsive
        response["success"] = False
        response["error"] = "API unavailable"
    except sr.UnknownValueError:
        # speech was unintelligible
        response["error"] = "Unable to recognize speech"

    return response