# Python 3 Script: Weather-Gatherer

# Purpose: Just testing out some API data retrieval (From https://home.openweathermap)
# and JSON file manipulation that's all.

# Date: 24/06/2020 Author: David Edwards

# Necessary modules for script
import pandas as pd
import time
import requests
from datetime import timedelta

# Personal openweathermap API key
api_key = 'f376dd072905eba93ba347dd7f54a0c2'

# master_list to hold all data for export to csv or json
master_list = []

# Put list of cities, countries, logitudes, etc into a dataframe
df = pd.read_excel('worldcities.xlsx')

# Stripping out relevant fields
dfw = df.filter(df[['city', 'country', 'lat', 'lng', 'iso2', 'iso3']])

# Looping through dfw dataframe and openweathermap API endpoint to retrieve weather for a bunch of countries
counter = 0
max_counter = 900
start_time = time.monotonic()

for k, v in dfw.iterrows():
    if counter < max_counter:
        country = v[1]
        city = v[0]
        lat = v[2]
        lng = v[3]
        url = 'https://api.openweathermap.org/data/2.5/onecall?lat=' + str(lat) + \
              '&lon=' + str(lng) + '&units=metric&exclude=hourly,daily&appid=' + api_key
        r = requests.get(url)
        if r.status_code == 200:
            data = r.json()
            timezone = data['timezone']
            sunrise = data['current']['sunrise']
            sunrise = pd.to_datetime(sunrise, unit='ms')
            sunset = data['current']['sunset']
            sunset = pd.to_datetime(sunset, unit='ms')
            temperature = data['current']['temp']
            feels_like = data['current']['feels_like']
            pressure = data['current']['pressure']
            humidity = data['current']['humidity']
            dew_point = data['current']['dew_point']
            ultra_violet_index = data['current']['uvi']
            clouds = data['current']['clouds']
            try:
                visibility = data['current']['visibility']
            except KeyError:
                visibility = 0
            wind_speed = data['current']['wind_speed']
            wind_direction = data['current']['wind_deg']
            current_weather = data['current']['weather'][0]['main']
            current_weather_desc = data['current']['weather'][0]['description']
            Country = {'country': country, 'city': city, 'timezone': timezone, 'current_weather': current_weather,
                       'current_weather_description': current_weather_desc, 'sunrise': sunrise, 'sunset': sunset,
                       'temperature_celsius': temperature, 'feels_like_celsius': feels_like,
                       'sea_level_atmospheric_pressure': pressure, 'humidity_%': humidity, 'dew_point': dew_point,
                       'ultra_violet_index': ultra_violet_index, 'clouds': clouds, 'visibility': visibility,
                       'wind_speed': wind_speed, 'wind_direction': wind_direction}
            master_list.append(Country)
            print(country, city, '|counter: ', counter)
        #  time.sleep(1)
        counter += 1

weather_data_expanded = pd.DataFrame(master_list)
print(weather_data_expanded.head())
weather_data_expanded.to_csv('weather_data_expanded.csv', index=False)
weather_data_expanded.to_json('weather_data_expanded.json')

end_time = time.monotonic()
print('Time taken: ', timedelta(seconds=end_time - start_time))

