# weather app
while True:
    import requests
    import json
    import win32com.client

    a  = win32com.client.Dispatch("SAPI.SpVoice")


    a.speak("Enter a the name of city ? ")
    city = input("Enter a the name of city ? ")

    url = f"https://api.weatherapi.com/v1/current.json?key=b67c70d4681b4aabaef170642251808&q={city}"
    r = requests.get(url)
    jason = json.loads(r.text)
   
    if(jason['error']['code']== 1006):
        print("Please enter a valid city name")
        a.speak("Please enter a valid city name")
        continue
    
    print('1 TO CLICK CHECK TEMPRATURE .')
    print('2 TO CLICK CHECK HUMIDITY')
    print(f"3. TO CLICK CHECK WIND DIRECTION in {city}")
    print(f"4 TO CLICK CHECK WIND SPEED(KPH)")

    a.speak("please choose an option ")
    choice = int(input("enter the choice number : "))

    if(choice==1):
        print(f"{jason["current"]["temp_c"]} celcius is temp in {city}")
        a.speak(f"{jason["current"]["temp_c"]} celcius is temp in {city}")
        
    elif(choice==2):
        print(f"the humidity in {city} is {jason["current"]["humidity"]} ")
        a.speak(f"the humidity in {city} is {jason['current']['humidity']} ")
    elif(choice==3):
        print(f"the wind direction  in {city} is {jason["current"]["wind_dir"]} ")
        a.speak(f"the wind direction  in {city} is {jason['current']['wind_dir']} ")
    else:
        print(f"the wind speed in {city} is {jason["current"]["wind_kph"]} kph ")
        a.speak(f"the wind speed in {city} is {jason['current']['wind_kph']} kph ")