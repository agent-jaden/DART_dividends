#-*- coding:utf-8 -*-
# Read urls from Excel file
# Parse information from itooza
# Write information to Excel file
import xlrd
import xlsxwriter
import os
from bs4 import BeautifulSoup
import urllib.request
import pickle
import getopt
import sys

def main():

	input_mode = 0
	try:
		opts, args = getopt.getopt(sys.argv[1:], "h")
	except getopt.GetoptError as err:
		print(err)
		sys.exit(2)
	for option, argument in opts:
		if option == "-h":
			help_msg = """
	Update DPS data from DART postings
			"""
			print(help_msg)
			sys.exit(2)
	
	### PART 1 - Read DPS Excel file
	num_stock = 2040
	input_file = "div_crawling_result.xlsx"
	cur_dir = os.getcwd()
	workbook_name = input_file
	workbook = xlrd.open_workbook(os.path.join(cur_dir, workbook_name))
	sheet_list = workbook.sheets()
	# DPS sheet
	sheet1 = sheet_list[1]

	stock_cat_list = []
	stock_name_list = []
	stock_num_list = []
	stock_url_list = []
	stock_dps_list = []
	name_error_list = []

	#for i in range(sheet1.row_len(0)):
	for i in range(num_stock):
		stock_cat_list.append(sheet1.cell(i+1,0).value)
		stock_name_list.append(sheet1.cell(i+1,1).value.strip())
		stock_num_list.append(int(sheet1.cell(i+1,2).value))
		stock_url_list.append(sheet1.cell(i+1,3).value)

		stock_dps_sub_list = []
		if sheet1.cell(i+1,4).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,4).value.replace(',','')))
		if sheet1.cell(i+1,5).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,5).value.replace(',','')))
		if sheet1.cell(i+1,6).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,6).value.replace(',','')))
		if sheet1.cell(i+1,7).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,7).value.replace(',','')))
		if sheet1.cell(i+1,8).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,8).value.replace(',','')))
		if sheet1.cell(i+1,9).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,9).value.replace(',','')))
		if sheet1.cell(i+1,10).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,10).value.replace(',','')))
		if sheet1.cell(i+1,11).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,11).value.replace(',','')))
		if sheet1.cell(i+1,12).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,12).value.replace(',','')))
		if sheet1.cell(i+1,13).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,13).value.replace(',','')))
		if sheet1.cell(i+1,14).value == 'N/A':
			stock_dps_sub_list.append(0.0)
		else:
			stock_dps_sub_list.append(float(sheet1.cell(i+1,14).value.replace(',','')))
		# New dps
		stock_dps_sub_list.append(0.0)
		stock_dps_list.append(stock_dps_sub_list)

	#print(stock_name_list)

	### PART 2 - Read DART posting Excel file
	input_file = "DART_dividends.xlsx"
	workbook_name = input_file
	workbook = xlrd.open_workbook(os.path.join(cur_dir, workbook_name))
	sheet_list = workbook.sheets()
	sheet1 = sheet_list[0]
	print(sheet1.nrows)

	for i in range(sheet1.nrows-1):
	
		#print(sheet1.cell(i+1,1).value)
		#배당구분 = 결산배당 & 배당종류 = 현금배당
		if ((sheet1.cell(i,5).value.strip() == "결산배당") or (sheet1.cell(i,5).value.strip() == "중간배당") or (sheet1.cell(i,5).value.strip() == "분기배당")) and (sheet1.cell(i,6).value.strip() == "현금배당"):
			try:
				find_index = stock_name_list.index(sheet1.cell(i+1,1).value.strip())
				if find_index != -1:
					#stock_dps_list[find_index][11] = stock_dps_list[find_index][11] + float(sheet1.cell(i+1,7).value.strip().replace(",",""))
					stock_dps_list[find_index][11] = float(sheet1.cell(i+1,7).value.strip().replace(",",""))
			except:
				name_error_list.append(sheet1.cell(i+1,1).value)
		
	print(name_error_list)

	### PART 3 - Write update DPS data
	workbook_name = "update_dps_result.xlsx"
	if os.path.isfile(os.path.join(cur_dir, workbook_name)):
		os.remove(os.path.join(cur_dir, workbook_name))
	workbook = xlsxwriter.Workbook(workbook_name)
	worksheet_dps = workbook.add_worksheet('DPS')

	filter_format = workbook.add_format({'bold':True,
										'fg_color': '#D7E4BC'
										})

	worksheet_dps.write(0, 0, "Category", filter_format)
	worksheet_dps.set_column('A:A', 15)
	worksheet_dps.write(0, 1, "Name", filter_format)
	worksheet_dps.set_column('B:B', 15)
	worksheet_dps.write(0, 2, "Code", filter_format)
	worksheet_dps.set_column('C:C', 10)
	worksheet_dps.write(0, 3, "URL", filter_format)
	worksheet_dps.set_column('D:D', 30)
	worksheet_dps.write(0, 4,  "2006.12", filter_format)
	worksheet_dps.write(0, 5,  "2007.12", filter_format)
	worksheet_dps.write(0, 6,  "2008.12", filter_format)
	worksheet_dps.write(0, 7,  "2009.12", filter_format)
	worksheet_dps.write(0, 8,  "2010.12", filter_format)
	worksheet_dps.write(0, 9,  "2011.12", filter_format)
	worksheet_dps.write(0, 10, "2012.12", filter_format)
	worksheet_dps.write(0, 11, "2013.12", filter_format)
	worksheet_dps.write(0, 12, "2014.12", filter_format)
	worksheet_dps.write(0, 13, "2015.12", filter_format)
	worksheet_dps.write(0, 14, "2016.12", filter_format)
	worksheet_dps.write(0, 15, "2017.12", filter_format)

	for k in range(num_stock):
		worksheet_dps.write(1+k, 0, stock_cat_list[k])
		worksheet_dps.write(1+k, 1, stock_name_list[k])
		worksheet_dps.write(1+k, 2, stock_num_list[k])
		worksheet_dps.write(1+k, 3, stock_url_list[k])
		for l in range(len(stock_dps_list[k])):
			worksheet_dps.write(1+k, 4+l, stock_dps_list[k][l])

	workbook.close()

# Main
if __name__ == "__main__":
	main()


