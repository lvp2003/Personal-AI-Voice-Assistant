import speech_recognition as sr
import os
import pyttsx3
import webbrowser
import win32com.client  # For Windows-specific TTS
import openai
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")  # For Windows

# Function to speak text
def say(text):
  speaker.Speak(text)

# Voice input function
def take_command():
  r = sr.Recognizer()
  with sr.Microphone() as source:
    print("Listening...")
    try:
      audio = r.listen(source)
      text = r.recognize_google(audio, language='en-in')
      print(f"You said: {text}")
      return text
    except sr.UnknownValueError:
      print("Could not understand audio")
      return None
    except sr.RequestError as e:
      print(f"Could not request results from Google Speech Recognition service; {e}")
      return None

# Function to generate chatbot response
def generate_chatbot_response(text):
  openai.api_key = "API_KEY"
  response = openai.Completion.create(
    engine="text-davinci-003-001",  # Update with the new model engine
    prompt=text,
    max_tokens=1024,
    n=1,
    stop=None,
    temperature=0.7,
)
  return response.choices[0].text.strip()

# Main Script
if __name__ == '__main__':
  say("Hello, I am Jarvis AI. How can I help you?")
  while True:
    text = take_command()
    if not text:
      continue

    # Exit command
    if 'close' in text.lower():
      say("Bye Sir, see you soon!")
      break

    # Predefined websites
    sites = [
      ['YouTube', 'https://www.youtube.com/'],
      ['Wikipedia', 'https://www.wikipedia.com/'],
      ['Google', 'https://www.google.com/'],
      ['Spotify', 'https://www.spotify.com/'],
      ['Facebook', 'https://www.facebook.com/']
    ]

    # Check for website commands
    for site in sites:
      if f"open {site[0].lower()}" in text.lower():
        say(f"Opening {site[0]}")
        webbrowser.open(site[1])

    # Time command
    if 'time' in text.lower():
      current_time = datetime.datetime.now().strftime('%I:%M %p')
      say(f"Sir, the current time is {current_time}")

    # Date command
    if 'date' in text.lower():
      current_date = datetime.datetime.now().strftime('%d/%m/%y')
      say(f"Sir, the current date is {current_date}")

    # Chatbot functionality
    if text.lower().startswith("jarvis,") or text.lower().startswith("hey jarvis"):
      print("chat with me")
      chatbot_response = generate_chatbot_response(text)
      say(chatbot_response)