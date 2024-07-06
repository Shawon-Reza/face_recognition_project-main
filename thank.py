import win32com.client

# Create a SpVoice object
speaker = win32com.client.Dispatch("SAPI.SpVoice") 

# Get the available voices
voices = speaker.GetVoices()

# Find the voice you want (e.g., "Microsoft David Desktop - English (United States)")
for voice in voices:
    if "Microsoft David Desktop" in voice.GetDescription():
        speaker.Voice = voice
        break

# Speak some text
speaker.Speak("Good morning everyone.I hope you all are doing well. My name is Shawon Reza, and i am here with my team member saiful islam rimon. Together, we are exited to present our project, TheReal Time Face Recognition and Attendance System. Throughout this presentation, we will walk you through the key features of our system, the technologies we used, and the benefits it offers over traditional attendance methods. We hope you find our work both interesting and insightful.Thank you for your attention, and let's get started.")

