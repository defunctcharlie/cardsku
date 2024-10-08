from random import randint
import pandas as pd
import sys
import os
import win32com.client
os.chdir(sys.path[0])

df = pd.read_excel('test.xlsm')

ExcelApp = win32com.client.GetActiveObject("Excel.Application")


cost = df['Cost']
SKU = df['SKU']


alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
		 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

sku_random = ''
cost_break = []
cost_output = ''
cost_input = str(cost)

def generate(cost_input):
	part_1 = ''
	part_2 = ''
	part_3 = ''
	
	#Randomly generate first part of SKU
	part_1 = str(randint(1, 1000))
	
	#Iterate through alphabet and flip a weighted coin to append letter to second part of SKU
	for i in alpha:
		rand_key = False
		if randint(1, 10) >= 9	:
			rand_key = True
		if rand_key == True:
			part_2 += i
	part_2 = str(part_2)
	cost_break = []
	#Create array based on cost input
	for i in cost_input:
		cost_break.append(i)
	
	#Iterate through array and create part 3 of SKU based on cypher
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
		part_3 += i

	sku_random = part_1 + "-" + part_2 + "-" + part_3
	return sku_random

count = 2
for i in cost:
	i = int(i)
	print(i)
	cell = 'I'
	ExcelApp.Range(cell + str(count)).Value = generate(str(i))
	count += 1
