import requests
import json
import win32com.client as wincom

speak = wincom.Dispatch("SAPI.SpVoice")
city = input("Enter name of the city: ")
url = f"https://api.weatherapi.com/v1/current.json?key=07890a163bfd475bbbe151432241912&q={city}"
response = requests.get(url)
weather_data = json.loads(response.text)

temperature = weather_data["current"]["temp_c"]
last_updated = weather_data["current"]["last_updated"]

print(f"\nThe temperature in {city} is {temperature}°C")
print(f"\nLast recorded time: {last_updated}")

speak.Speak(f"The temperature in {city} is {temperature} degrees Celsius")
speak.Speak(f"The data was last updated on {last_updated}.")
