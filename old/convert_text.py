import pandas as pd
from datetime import date
import os

# order_name = 'Kim Seng KK6 外送 31_3_2022 (Responses) - Form Responses 1.csv'
input_path = os.path.join('input',"filename.txt")

f = open(input_path,"r", encoding='utf-8')
order_name = f.readlines()[0] 

order = pd.read_csv(os.path.join('input',order_name)  +'.csv')

# print(order)
# print(order.columns)

dishes = order[order.columns[5]].tolist()
customizes = order[order.columns[6]].fillna('No Customizes').tolist()

#write text
today_ = date.today().strftime("%d.%m.%Y")
txt_name = f"{today_} orders.txt"
file_path = os.path.join('output', txt_name)

if not os.path.exists('output'):
	os.makedirs('output')

with open(file_path, 'w', encoding="utf-8") as writer:
	for index, (d, c) in enumerate(zip(dishes, customizes)):
		writer.write(f'{index+1}. \n')
		for dish in d.split(','):
			dish = dish.strip()
			writer.write(dish)
			writer.write('\n')

		for customize in c.split(','):
			customize = customize.strip()
			writer.write(customize + '\n')

		writer.write('\n')


