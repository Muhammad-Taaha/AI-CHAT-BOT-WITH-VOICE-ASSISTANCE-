import win32com.client
import speech_recognition as sr
import webbrowser
import datetime
from gpt import text_manager
from colorama import init, Fore, Back, Style

# Initialize colorama
init(autoreset=True)

# Initialize speaker for text-to-speech
speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    """Speak out the given text."""
    speaker.Speak(text)

def take_command():
    """Listen to the user's voice command and return the recognized text."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print(Fore.CYAN + Style.BRIGHT + "Listening...")  # Cyan text for listening
        r.pause_threshold = 1
        try:
            audio = r.listen(source)
            query = r.recognize_google(audio, language="en-in")
            print(Fore.GREEN + Style.BRIGHT + f"User said: {query}")  # Green text for recognized speech
            return query.lower().strip()
        except Exception as e:
            print(Fore.RED + Style.BRIGHT + "Error:", e)  # Red text for error
            return "I didn't catch that. Please try again."

def main():
    print(Fore.YELLOW + Back.BLUE + Style.BRIGHT + "What do you want to do?")  # Yellow text on blue background
    print(Fore.YELLOW + Style.BRIGHT + "Enter the index:")
    print(Fore.GREEN + Style.BRIGHT + "1: GIVE TEXT INPUT")
    print(Fore.GREEN + Style.BRIGHT + "2: GIVE VOICE COMMAND")

    var_2 = int(input(Fore.MAGENTA + Style.BRIGHT + "ENTER THE CHOICE: "))  # Magenta input prompt

    if var_2 == 1:
        while True:
            var_1 = input(Fore.CYAN + Back.BLACK + Style.BRIGHT + "Enter the task: ")  # Cyan text, black background
            if var_1 in ['stop', 'end', 'exit']:
                break
            print(Fore.YELLOW + Back.BLACK +Style.BRIGHT+ f"Performing: {var_1}")  # Green bold output
            print(text_manager(var_1))

    elif var_2 == 2:
        print(Fore.YELLOW + Back.BLACK + Style.BRIGHT + "I AM A CHATBOT CREATED BY TAAHA")
        say("Welcome! How can I help you?")

        while True:
            query = take_command()
            if "exit" in query or "stop" in query:
                print(Fore.YELLOW + Back.GREEN + Style.BRIGHT + "Goodbye!")  # Yellow text, green background for exit
                say("Goodbye!")
                break

            sites = [
    # Search Engines
                    ['google', 'https://google.com'],
                    ['bing', 'https://bing.com'],
                    ['duckduckgo', 'https://duckduckgo.com'],
                    ['yahoo', 'https://yahoo.com'],
                    ['baidu', 'https://baidu.com'],
                    ['yandex', 'https://yandex.com'],
                    
                    # Social Media
                    ['facebook', 'https://facebook.com'],
                    ['instagram', 'https://instagram.com'],
                    ['twitter', 'https://twitter.com'],
                    ['linkedin', 'https://linkedin.com'],
                    ['pinterest', 'https://pinterest.com'],
                    ['reddit', 'https://reddit.com'],
                    ['tiktok', 'https://www.tiktok.com'],
                    ['snapchat', 'https://www.snapchat.com'],
                    ['tumblr', 'https://www.tumblr.com'],
                    ['whatsapp', 'https://web.whatsapp.com'],
                    
                    # Video Platforms
                    ['youtube', 'https://youtube.com'],
                    ['vimeo', 'https://vimeo.com'],
                    ['twitch', 'https://twitch.tv'],
                    ['dailymotion', 'https://www.dailymotion.com'],
                    ['netflix', 'https://netflix.com'],
                    ['hulu', 'https://www.hulu.com'],
                    ['disney+', 'https://www.disneyplus.com'],
                    ['prime video', 'https://primevideo.com'],
                    ['peacock', 'https://www.peacocktv.com'],
                    
                    # Music & Entertainment
                    ['spotify', 'https://spotify.com'],
                    ['soundcloud', 'https://soundcloud.com'],
                    ['apple music', 'https://music.apple.com'],
                    ['pandora', 'https://www.pandora.com'],
                    ['tidal', 'https://www.tidal.com'],
                    ['deezer', 'https://www.deezer.com'],
                    
                    # Online Shopping
                    ['amazon', 'https://amazon.com'],
                    ['flipkart', 'https://flipkart.com'],
                    ['ebay', 'https://ebay.com'],
                    ['etsy', 'https://etsy.com'],
                    ['walmart', 'https://walmart.com'],
                    ['alibaba', 'https://www.alibaba.com'],
                    ['shein', 'https://www.shein.com'],
                    ['zara', 'https://www.zara.com'],
                    ['bestbuy', 'https://www.bestbuy.com'],
                    ['target', 'https://www.target.com'],
                    
                    # Technology & Development
                    ['github', 'https://github.com'],
                    ['gitlab', 'https://gitlab.com'],
                    ['stackoverflow', 'https://stackoverflow.com'],
                    ['npm', 'https://www.npmjs.com'],
                    ['python docs', 'https://docs.python.org'],
                    ['mozilla developer', 'https://developer.mozilla.org'],
                    ['w3schools', 'https://www.w3schools.com'],
                    ['hackerrank', 'https://www.hackerrank.com'],
                    ['freecodecamp', 'https://www.freecodecamp.org'],
                    
                    # News & Information
                    ['cnn', 'https://cnn.com'],
                    ['bbc', 'https://bbc.com'],
                    ['nytimes', 'https://nytimes.com'],
                    ['the guardian', 'https://www.theguardian.com'],
                    ['al jazeera', 'https://aljazeera.com'],
                    ['ndtv', 'https://ndtv.com'],
                    ['reuters', 'https://www.reuters.com'],
                    ['forbes', 'https://www.forbes.com'],
                    ['huffpost', 'https://www.huffpost.com'],
                    ['the verge', 'https://www.theverge.com'],
                    
                    # Education
                    ['wikipedia', 'https://wikipedia.org'],
                    ['khan academy', 'https://www.khanacademy.org'],
                    ['coursera', 'https://www.coursera.org'],
                    ['udemy', 'https://www.udemy.com'],
                    ['edx', 'https://www.edx.org'],
                    ['duolingo', 'https://www.duolingo.com'],
                    ['quizlet', 'https://quizlet.com'],
                    ['codeacademy', 'https://www.codecademy.com'],
                    ['duolingo', 'https://www.duolingo.com'],
                    ['learncpp', 'https://www.learncpp.com'],
                    
                    # Food Delivery
                    ['zomato', 'https://zomato.com'],
                    ['swiggy', 'https://swiggy.com'],
                    ['uber eats', 'https://www.ubereats.com'],
                    ['grubhub', 'https://www.grubhub.com'],
                    ['foodpanda', 'https://www.foodpanda.com'],
                    ['door dash', 'https://www.doordash.com'],
                    
                    # Productivity & Tools
                    ['gmail', 'https://mail.google.com'],
                    ['outlook', 'https://outlook.live.com'],
                    ['drive', 'https://drive.google.com'],
                    ['dropbox', 'https://dropbox.com'],
                    ['onedrive', 'https://onedrive.live.com'],
                    ['trello', 'https://trello.com'],
                    ['notion', 'https://notion.so'],
                    ['slack', 'https://slack.com'],
                    ['zoom', 'https://zoom.us'],
                    ['skype', 'https://www.skype.com'],
                    ['todoist', 'https://todoist.com'],
                    
                    # Travel & Maps
                    ['google maps', 'https://maps.google.com'],
                    ['tripadvisor', 'https://www.tripadvisor.com'],
                    ['airbnb', 'https://www.airbnb.com'],
                    ['booking', 'https://www.booking.com'],
                    ['expedia', 'https://www.expedia.com'],
                    ['makemytrip', 'https://www.makemytrip.com'],
                    ['uber', 'https://www.uber.com'],
                    ['lyft', 'https://www.lyft.com'],
                    ['skyscanner', 'https://www.skyscanner.com'],
                    
                    # Weather
                    ['weather', 'https://weather.com'],
                    ['accuweather', 'https://www.accuweather.com'],
                    ['bbc weather', 'https://www.bbc.com/weather'],
                    ['met office', 'https://www.metoffice.gov.uk'],
                    
                    # Miscellaneous
                    ['quora', 'https://quora.com'],
                    ['medium', 'https://medium.com'],
                    ['paypal', 'https://paypal.com'],
                    ['adobe', 'https://adobe.com'],
                    ['canva', 'https://canva.com'],
                    ['bitly', 'https://bitly.com'],
                    ['speedtest', 'https://www.speedtest.net'],
                    ['hightail', 'https://www.hightail.com'],
                    
                    # Desktop Applications
                    ['workbench', r'C:\Program Files\MySQL\MySQL Workbench 8.0 CE\MySQLWorkbench.exe'],
                    ['excel', r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"],
                    ['access', r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"],
                    ['powerpoint', r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"],
                    ['word', r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"],
                    ['visual studio code', r"C:\Program Files\Microsoft VS Code\Code.exe"],
                    ['notepad ', r"C:\Program Files\Notepad++\notepad++.exe"]
                ]


            if "open" in query:
                found = False
                for site in sites:
                    if site[0] in query:
                        print(Fore.YELLOW + Back.GREEN + f"Opening {site[0]}...")  # Yellow text on green background
                        webbrowser.open(site[1])
                        found = True
                        break
                if not found:
                    print(Fore.RED + Back.WHITE + Style.BRIGHT + "Sorry, I don't recognize that site.")  # Red text on white background for error

            elif "time" in query:
                time = datetime.datetime.now().strftime('%H:%M:%S')
                print(Fore.CYAN + Back.BLACK + Style.BRIGHT + f"The time is: {time}")  # Cyan text with black background
                say(f"The time is {time}")
            else:
                print(Fore.RED + Back.WHITE + Style.BRIGHT + "I didn't understand that. Please try again.")  # Red text on white background for unrecognized input

if __name__ == "__main__":
    main()

