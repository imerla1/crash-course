import win32com.client

# Calling the Disptach method of the module which  
# interact with Microsoft Speech SDK to speak 
# the given input from the keyboard 

speaker = win32com.client.Dispatch("SAPI.SpVoice")
s = 'hello world'
speaker.Speak(s)
