import win32com.client


# Parameters for voice

# Text to Speech initialization
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# range 0(low) - 100(loud)
volume = 100

# range -10(slow)  +10(fast)
rate = 1

speaker.Rate = rate
speaker.Volume = volume


# Say Func
def say(text):
    speaker.Speak(text)


# npm install win32com.client