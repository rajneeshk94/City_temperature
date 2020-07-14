import requests
import openpyxl
import sys

run = True 

#Gets temperature of city in F
def get_temp(city):
	url = f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid=1694fcee8683b88cc0f0dc7845f7087e'
	res = requests.get(url)
	city_data = res.json()
	temp = city_data['main']['temp']
	return temp

#Gets humidity data
def get_humi(city):
	url = f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid=1694fcee8683b88cc0f0dc7845f7087e'
	res = requests.get(url)
	city_data = res.json()
	humidity = city_data['main']['humidity']
	return humidity

def update_data():
	
	#loads workbook and the Weather worksheet
	wb = openpyxl.load_workbook('City_temperatures.xlsx')
	weather = wb['Weather']

	#Updates temperature and humidity city wise and unit wise
	for city,unit,temperature,humidity,update in zip(weather['A'], weather['D'], weather['B'], weather['C'], weather['E']):
		#Skips the first two column
		if city.value == 'City Name' or city.value == None:
			continue
		#Checks Unit and Update value
		elif unit.value == 'C' and update.value == 1:
			temperature.value = round((get_temp(city.value) - 32)*(5/9), 0) 
		
		elif unit.value == 'F' and update.value == 1:
			temperature.value = round(get_temp(city.value), 0)
		
		#Checks update value for humidity and updates it
		if update.value == 1:
			humidity.value = get_humi(city.value)
		else:
			pass

	wb.save('City_temperatures.xlsx')			
	print('Updated..')

try:
	while run:
		update_data()

except KeyboardInterrupt:
	sys.exit()