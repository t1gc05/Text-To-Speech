Dim message
message = InputBox("Type your message below:", "Text To Speech")
Set sapi = CreateObject("SAPI.SpVoice")
sapi.Speak message