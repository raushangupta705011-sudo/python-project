import requests
import datetime
import pandas as pd

def get_weather_data(location, api_key):
    url = f"http://api.openweathermap.org/data/2.5/forecast?q={location}&appid={api_key}&units=metric"
    response = requests.get(url)
    return response.json()

def is_suitable_for_farming(temp, rain_chances):
    if temp > 0 and temp < 30 and rain_chances == 'No':
        return "Suitable for farming"
    else:
        return "Not suitable for farming"

def display_weather_forecast(location, api_key):
    data = get_weather_data(location, api_key)
    weather_records = []  # List to hold the weather records

    if 'list' in data:
        print(f"\n7-Day Farming Weather Forecast for {location}:\n")
        print(f"{'Date':<15} {'Temp (°C)':<12} {'Humidity (%)':<15} {'Weather':<20} {'Farming Advice':<25}")
        print('-' * 90)

        for i in range(0, len(data['list']), 8):  # Taking one reading per day
            day = data['list'][i]
            date = datetime.datetime.fromtimestamp(day['dt']).strftime('%Y-%m-%d')
            temp = day['main']['temp']
            humidity = day['main']['humidity']
            weather = day['weather'][0]['main']
            rain_chances = 'Yes' if 'rain' in weather.lower() else 'No'
            farming_advice = is_suitable_for_farming(temp, rain_chances)

            print(f"{date:<15} {temp:<12} {humidity:<15} {weather:<20} {farming_advice:<25}")

            # Append each day's data to the list
            weather_records.append({
                "Date": date,
                "Temperature (°C)": temp,
                "Humidity (%)": humidity,
                "Weather": weather,
                "Farming Advice": farming_advice
            })

        print('-' * 90)

        # Create a DataFrame and save it as an Excel file
        df = pd.DataFrame(weather_records)
        excel_filename = f"{location}_weather_forecast.xlsx"
        df.to_excel(excel_filename, index=False)
        print(f"Weather data saved to {excel_filename}")
    else:
        print("Weather data not found. Please check the location name and API key.")

if __name__ == "__main__":
    api_key = ""  # Replace with your actual OpenWeatherMap API key
    location = input("Enter a city or village name: ")
    display_weather_forecast(location, api_key)
