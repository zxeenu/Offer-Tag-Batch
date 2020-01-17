import openpyxl
import os
from tkinter import *
from tkinter.filedialog import askopenfilename

class Tag_Data:

	def __init__(self, name, normal_price, offer_price, packing, expiry):
		self.name = name
		self.normal_price = normal_price
		self.offer_price = offer_price
		self.packing = packing
		self.expiry = expiry

def excel_finder():
	root = Tk()
	root.withdraw()
	path_to_file = askopenfilename()
	root.deiconify()
	root.destroy()
	root.quit()
	return path_to_file

# for testing
os.chdir(r"D:\PROJECTS - OTHER\work_scripts\offer_tag_maker\batch")
workbook = openpyxl.load_workbook("EXPIRY  2020.xlsx")
sheet = workbook["JANUARY"]

column_name = 1
column_packing = 2
column_expiry = 3
column_retail_price = 6
column_offer_price = 7

# here ends for testing

# manual dynamic selection for actual usage
# path = excel_finder()
# workbook = openpyxl.load_workbook(path)
# sheet_name = input("Please enter the name of the sheet\n")
# sheet = workbook[sheet_name]

# column_name = int(input("Column number for the name\n"))
# column_packing = int(input("Column number for the packing\n"))
# column_expiry = int(input("Column number for the expiry\n"))
# column_retail_price = int(input("Column number for the retail price\n"))
# column_offer_price = int(input("Column number for the offer price\n"))

# column_name = column_name - 1
# column_packing = column_packing - 1
# column_expiry = column_expiry - 1
# column_retail_price = column_retail_price - 1
# column_offer_price = column_offer_price - 1

# here ends code for manual action

tag_list = []

# while True:
# 	name = sheet.cell(row=current_row, column=column_name).value
# 	packing = sheet.cell(row=current_row, column=column_packing).value
# 	expiry = sheet.cell(row=current_row, column=column_expiry).value
# 	retail_price = sheet.cell(row=current_row, column=column_retail_price).value
# 	offer_price = sheet.cell(row=current_row, column=column_offer_price).value

# 	tag_obj = Tag_Data(name, retail_price, offer_price, packing, expiry)
# 	tag_list.append(tag_obj)


# 	current_row = current_row+1
# 	if name == None:
# 		break
	
# for tag in tag_list:
# 	print(f"{tag.name} - {tag.normal_price} - {tag.offer_price} - {tag.packing} - {tag.expiry}")


##########################

# for row in sheet.iter_rows(values_only=True):
# 	name = sheet.cell(row=current_row, column=column_name).value
# 	packing = sheet.cell(row=current_row, column=column_packing).value
# 	expiry = sheet.cell(row=current_row, column=column_expiry).value
# 	retail_price = sheet.cell(row=current_row, column=column_retail_price).value
# 	offer_price = sheet.cell(row=current_row, column=column_offer_price).value

# 	tag_obj = Tag_Data(name, retail_price, offer_price, packing, expiry)
# 	tag_list.append(tag_obj)

# 	print(row)
# 	current_row = current_row+1

# for tag in tag_list:
# 	print(f"{tag.name} - {tag.normal_price} - {tag.offer_price} - {tag.packing} - {tag.expiry}")

for row in sheet.iter_rows(values_only=True):
	name = row[column_name]
	packing = row[column_packing]
	expiry = row[column_expiry]
	retail_price = row[column_retail_price]
	offer_price = row[column_offer_price]

	tag_obj = Tag_Data(name, retail_price, offer_price, packing, expiry)
	tag_list.append(tag_obj)

for tag in tag_list:
	print(f"{tag.name} - {tag.normal_price} - {tag.offer_price} - {tag.packing} - {tag.expiry}")