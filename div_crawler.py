#-*- coding:utf-8 -*-
# Read urls from Excel file
# Parse information of dividends from itooza (www.itooza.com)
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
		opts, args = getopt.getopt(sys.argv[1:], "h", ["mode="])
	except getopt.GetoptError as err:
		print(err)
		sys.exit(2)
	for option, argument in opts:
		if option == "-h":
			help_msg = """
	input mode 0: Crawling 1: Pickle
			"""
			print(help_msg)
			sys.exit(2)
		elif option == "--mode":
			input_mode = int(argument)

	### PART I - Read Excel file for stock lists
	num_stock = 2046
	#num_stock = 100
	input_file = "basic_20180309.xlsx"
	cur_dir = os.getcwd()
	workbook_name = input_file
	
	stock_cat_list = []
	stock_name_list = []
	stock_num_list = []
	stock_url_list = []
	
	workbook = xlrd.open_workbook(os.path.join(cur_dir, workbook_name))
	sheet_list = workbook.sheets()
	sheet1 = sheet_list[0]

	for i in range(num_stock):
		stock_cat_list.append(sheet1.cell(i+1,0).value)
		stock_name_list.append(sheet1.cell(i+1,1).value)
		stock_num_list.append(int(sheet1.cell(i+1,2).value))
		stock_url_list.append(sheet1.cell(i+1,3).value)

	### PART II - Read information from URLs
	dps_list = []
	dps_ratio_list = []
	error_list = []

	if input_mode == 0:
		for j in range(num_stock):
			print(j, stock_name_list[j])
			if j%10 == 0: print (j)
			
			try:
				dps_sub_list = []
				dps_ratio_sub_list = []
				
				url = stock_url_list[j]
				#print(url)
				
				handle = None
				while handle == None:
					try:
						handle = urllib.request.urlopen(url)
						#print(handle)
					except:
						pass

				data = handle.read()
				soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
				
				# Find table
				table = soup.findAll('div', {'id':'indexTable2'})
				
				#print(soup.prettify())
				#print(type(soup))
				#print(type(table))
		
				# TH list
				# [0] 투자지표
				# [1] 17.09월
				# [2] 16.12월
				# [3] 15.12월
				# [4] 14.12월
				# [5] 13.12월
				# [6] 12.12월
				# [7] 11.12월
				# [8] 10.12월
				# [9] 09.12월
				# [10] 08.12월
				# [11] 07.12월
				# [12] 06.12월

				th_list = table[0].findAll('th')
				#for ths  in th_list:
				#	print (ths.text)
				# tr 0  : Column
				# tr 1  : 주당순이익(EPS,연결지배)
				# tr 2  : 주당순이익(EPS,개별)
				# tr 3  : PER (배)
				# tr 4  : 주당순자산(지분법)
				# tr 5  : PBR (배)
				# tr 6  : 주당 배당금
				# tr 7  : 시가 배당률 (%)
				# tr 8  : ROE (%)
				# tr 9  : 순이익률 (%)
				# tr 10 : 영업이익률 (%)
				# tr 11 : 주가	
				tr_list = table[0].findAll('tr')
				
				# Get information for dividens 
				td_list = tr_list[6].findAll('td')
				for tds  in td_list:
					dps_sub_list.append(tds.text)

				td_list2 = tr_list[7].findAll('td')
				for tds2  in td_list2:
					dps_ratio_sub_list.append(tds2.text)
					#print(tds2.text)
			except:
				dps_sub_list = [0,0,0,0,0,0,0,0,0,0,0,0]
				dps_ratio_sub_list = [0,0,0,0,0,0,0,0,0,0,0,0]
				error_list.append(stock_name_list[j])

			dps_sub_list.reverse()
			dps_ratio_sub_list.reverse()
			dps_list.append(dps_sub_list)
			dps_ratio_list.append(dps_ratio_sub_list)


		f_pickle = open("crawling_list", "wb")
		pickle.dump(dps_list, f_pickle)
		pickle.dump(dps_ratio_list, f_pickle)
		f_pickle.close()

		print(error_list)
	
	# input mode is not 0
	else:
		f_pickle = open("crawling_list", "rb")
		dps_list = pickle.load(f_pickle)
		dps_ratio_list = pickle.load(f_pickle)

	### PART III - Write information to new Excel file
	workbook_name = "div_crawling_result.xlsx"
	if os.path.isfile(os.path.join(cur_dir, workbook_name)):
		os.remove(os.path.join(cur_dir, workbook_name))
	workbook = xlsxwriter.Workbook(workbook_name)
	worksheet_raw = workbook.add_worksheet('RAW')
	worksheet_dps = workbook.add_worksheet('DPS')
	worksheet_ratio = workbook.add_worksheet('RATIO')

	filter_format = workbook.add_format({'bold':True,
										'fg_color': '#D7E4BC'
										})
	# Write filter
	worksheet_raw.write(0, 0, "Category", filter_format)
	worksheet_raw.set_column('A:A', 15)
	worksheet_raw.write(0, 1, "Name", filter_format)
	worksheet_raw.set_column('B:B', 15)
	worksheet_raw.write(0, 2, "Code", filter_format)
	worksheet_raw.set_column('C:C', 10)
	worksheet_raw.write(0, 3, "URL", filter_format)
	worksheet_raw.set_column('D:D', 30)
	worksheet_raw.write(0, 4,  "2006.12", filter_format)
	worksheet_raw.write(0, 5,  "2007.12", filter_format)
	worksheet_raw.write(0, 6,  "2008.12", filter_format)
	worksheet_raw.write(0, 7,  "2009.12", filter_format)
	worksheet_raw.write(0, 8,  "2010.12", filter_format)
	worksheet_raw.write(0, 9,  "2011.12", filter_format)
	worksheet_raw.write(0, 10, "2012.12", filter_format)
	worksheet_raw.write(0, 11, "2013.12", filter_format)
	worksheet_raw.write(0, 12, "2014.12", filter_format)
	worksheet_raw.write(0, 13, "2015.12", filter_format)
	worksheet_raw.write(0, 14, "2016.12", filter_format)
	worksheet_raw.write(0, 15, "2017.12", filter_format)

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

	worksheet_ratio.write(0, 0, "Category", filter_format)
	worksheet_ratio.set_column('A:A', 15)
	worksheet_ratio.write(0, 1, "Name", filter_format)
	worksheet_ratio.set_column('B:B', 15)
	worksheet_ratio.write(0, 2, "Code", filter_format)
	worksheet_ratio.set_column('C:C', 10)
	worksheet_ratio.write(0, 3, "URL", filter_format)
	worksheet_ratio.set_column('D:D', 30)
	worksheet_ratio.write(0, 4,  "2006.12", filter_format)
	worksheet_ratio.write(0, 5,  "2007.12", filter_format)
	worksheet_ratio.write(0, 6,  "2008.12", filter_format)
	worksheet_ratio.write(0, 7,  "2009.12", filter_format)
	worksheet_ratio.write(0, 8,  "2010.12", filter_format)
	worksheet_ratio.write(0, 9,  "2011.12", filter_format)
	worksheet_ratio.write(0, 10, "2012.12", filter_format)
	worksheet_ratio.write(0, 11, "2013.12", filter_format)
	worksheet_ratio.write(0, 12, "2014.12", filter_format)
	worksheet_ratio.write(0, 13, "2015.12", filter_format)
	worksheet_ratio.write(0, 14, "2016.12", filter_format)
	worksheet_ratio.write(0, 15, "2017.12", filter_format)

	for k in range(num_stock):
		worksheet_raw.write(1+k*2, 0, stock_cat_list[k])
		worksheet_raw.write(1+k*2, 1, stock_name_list[k])
		worksheet_raw.write(1+k*2, 2, stock_num_list[k])
		worksheet_raw.write(1+k*2, 3, stock_url_list[k])
		
		worksheet_dps.write(1+k, 0, stock_cat_list[k])
		worksheet_dps.write(1+k, 1, stock_name_list[k])
		worksheet_dps.write(1+k, 2, stock_num_list[k])
		worksheet_dps.write(1+k, 3, stock_url_list[k])

		worksheet_ratio.write(1+k, 0, stock_cat_list[k])
		worksheet_ratio.write(1+k, 1, stock_name_list[k])
		worksheet_ratio.write(1+k, 2, stock_num_list[k])
		worksheet_ratio.write(1+k, 3, stock_url_list[k])

		for l in range(len(dps_list[k])):
			worksheet_raw.write(1+k*2, 4+l, dps_list[k][l])
			worksheet_dps.write(1+k, 4+l, dps_list[k][l])

		for m in range(len(dps_ratio_list[k])):
			worksheet_raw.write(2+k*2, 4+m, dps_ratio_list[k][m])
			worksheet_ratio.write(1+k, 4+m, dps_ratio_list[k][m])

# Main
if __name__ == "__main__":
	main()


