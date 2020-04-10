from win32com.client import Dispatch

def speak(str):
	speak =  Dispatch(("SAPI.SpVoice"))
	speak.Speak(str)

if __name__ == '__main__':
	speak("I cannot believe another year has passed. It feels like we just met yesterday, but at the same time it feels like I have known you all my life. You make time meaningless. In fact, you make everything else feel meaningless because the only thing that matters is you. You have brought so much light into my life. I would be lost without your torch. Thank you for everything you have done for me — and thank you for helping me grow into the woman I have become.To my forever person,I love us. We’re the cutest. I know that sounds braggy, but I mean it when I say that I think we make the perfect couple. We understand each other. We listen to each other. We inspire each other to become stronger with each passing day. Happy anniversary. I cannot wait to spend another year alongside you, because there is no place I would rather be. You’re stuck with me. You better remember that!")