from datetime import date
from math import cos
from random import randint
import pandas as pd
import sys
import os
import win32com.client
os.chdir(sys.path[0])

df = pd.read_excel('test.xlsm')

ExcelApp = win32com.client.GetActiveObject("Excel.Application")

print(date.today())
cost = df['Cost per Item']
name = df['Name']
SKU = df['Barcode']
sport = df['Sport']
SKU_real = df['SKU']

alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
		 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

sku_random = ''
cost_break = []
cost_output = ''
cost_input = str(cost)
name_input = name
sport_input = sport
print(sport_input)
def generate(cost_input, name_input, sport_input):
	part_1 = ''
	part_2 = ''
	part_3 = ''
	part_4 = ''
	part_5 = ''
	
	#Test for Raw or Graded part 1
	if "PSA" in name_input:
		part_1 = 'G'
	elif "BGS" in name_input:
		part_1 = 'G'
	else:
		part_1 = 'R'
	
	#Code for sport part 2
	if "Football" in sport_input:
		part_2 = "F"
	elif "Basketball" in sport_input:
		part_2 = "BA"
	elif "Baseball" in sport_input:
		part_2 = "BB"
	elif "UFC" in sport_input:
		part_2 = "U"
	elif "Hockey" in sport_input:
		part_2 = "H"
	elif "Soccer" in sport_input:
		part_2 = "S"
	elif "Wrestling" in sport_input:
		part_2 = "W"
	elif "Misc" in sport_input:
		part_2 = "M"
	else:
		part_2 = "M"

	cost_break = []

	#Year and date code for part 3
	date_key = str(date.today())
	date_2 = date_key[2]
	date_3 = date_key[3]
	date_5 = date_key[5]
	date_6 = date_key[6]
	part_3 = date_2 + date_3 + date_5 + date_6


	#Create array based on cost input
	for i in cost_input:
		cost_break.append(i)
	
	#Iterate through array and create part 4 of SKU based on cypher
	pos_check = 0
	for i in cost_break:
		pos_check += 1
		i = int(i)
		if pos_check % 2 == 0:
			if i == 0:
				i = "A"
			elif i == 1:
				i = "B"
			elif i == 2:
				i = "C"
			elif i == 3:
				i = "D"
			elif i == 4:
				i = "E"
			elif i == 5:
				i = "F"
			elif i == 6:
				i = "G"
			elif i == 7:
				i = "H"
			elif i == 8:
				i = "I"
			elif i == 9:
				i = "J"
		i = str(i)	
		part_4 += i

	for i in alpha:
		rand_key = False
		if randint(0,100) >= 92	:
			rand_key = True
		if rand_key == True:
			part_5 += i
	part_5 += str(randint(1, 999))


	sku_random = part_1 + "-" + part_2 + "-" + part_3 + "-" + part_4 + "-" + part_5
	return sku_random

count = 2
erase_count = 1
name_count = 0
for i in range(500):
	erase_count += 1
	cell_2 = 'N'
	cell_3 = 'O'
	ExcelApp.Range(cell_2 + str(erase_count)).value = ""
	ExcelApp.Range(cell_3 + str(erase_count)).value = ""
for i in cost:
	i =		int(i)
	cell = 'I'
	cell_2 = 'N'
	cell_3 = 'O'
	ExcelApp.Range(cell + str(count)).Value = generate(str(i), name[name_count], sport[name_count])
	ExcelApp.Range(cell_2 + str(count)).Value = generate(str(i), name[name_count], sport[name_count])
	ExcelApp.Range(cell_3 + str(count)).Value = ExcelApp.Range(cell + str(count)).Value
	count += 1
	name_count += 1


