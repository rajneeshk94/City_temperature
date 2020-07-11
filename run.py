import requests
import openpyxl
import sys
import threading

city_names = ['London', 'New Delhi', 'New York', 'Singapore']
temps = []

def update_temp():
	threading.Timer(5.0, update_temp).start() #Updates temperature after every 5 seconds
	
	for city_name in city_names:
		url = f'https://api.openweathermap.org/data/2.5/weather?q={city_name}&appid=1694fcee8683b88cc0f0dc7845f7087e'
		res = requests.get(url)

		city_data = res.json()

		temp = city_data['main']['temp']
		temps.append(temp)
	# 	print(temp)
	# print('\n')

	wb = openpyxl.load_workbook('City_temperatures.xlsx')

	ws1 = wb['Sheet1']
	ws1['A1'] = 'City Name'
	ws1['B1'] = 'Temperature in F'
	ws1['C1'] = 'Temperature in C'

	for i in range(1, len(city_names) + 1):
		ws1[f'A{i+2}'] = city_names[i - 1]

	for i in range(1, len(city_names) + 1):
		ws1[f'B{i+2}'] = temps[i - 1]

	for i in range(1, len(city_names) + 1):
		ws1[f'C{i+2}'] = (temps[i - 1] - 32) * (5/9)			
	
	wb.save('City_temperatures.xlsx')


try:
	update_temp()

#Press CTRL + C to exit
except KeyboardInterrupt:
	sys.exit()