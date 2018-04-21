#-*- coding:utf-8 -*-
# Parsing dividends data from DART
import urllib.request
import urllib.parse
import xlsxwriter
import os
import time
import sys
import getopt
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import re

def main():

	# Default
	config_mode = 0
	config_start_year	= 2018
	config_start_month	= 4
	config_start_day	= 19
	config_end_year		= 2018
	config_end_month	= 4
	config_end_day		= 20
	corp = "삼성전자"
	workbook_name = "DART_insider_buy.xlsx"

	try:
		opts, args = getopt.getopt(sys.argv[1:], "m:s:e:c:o:h", ["mode=", "start=", "end=", "corp=", "output", "help"])
	except getopt.GetoptError as err:
		print(err)
		sys.exit(2)
	for option, argument in opts:
		if option == "-h" or option == "--help":
			help_msg = """
================================================================================
-m or --mode <number>   :  Operation Mode
                            0 : Find posting of dividends in specific period
                            1 : Find all posting of dividends for specific corporation
-s or --start <number>  :  Start of period
                            year(4digits) + month(2digits) + day(2digits)
-e or --end <number>    :  End of period
                            year(4digits) + month(2digits) + day(2digits)
-c or --corp <name>     :  Corporation name
-o or --output <name>	:  Output file name
-h or --help            :  Show help messages

<Example>
>> python dart_dividends.py -m 0 -s 20171115 -e 20171215 -o out_file_name
>> python dart_dividends.py -m 1 -c S-Oil
================================================================================
					"""
			print(help_msg)
			sys.exit(2)
		elif option == "--mode" or option == "-m":
			config_mode = int(argument)
		elif option == "--start" or option == "-s":
			config_start_year	= int(argument[0:4])
			config_start_month	= int(argument[4:6])
			config_start_day	= int(argument[6:8])
		elif option == "--end" or option == "-e":
			config_end_year		= int(argument[0:4])
			config_end_month	= int(argument[4:6])
			config_end_day		= int(argument[6:8])
		elif option == "--corp" or option == "-c":
			corp = argument
		elif option == "--output" or option == "-o":
			workbook_name = argument + ".xlsx"

	# URL for Mode 0
	url_templete_0 = "http://dart.fss.or.kr/dsab002/search.ax?reportName=%s&&maxResults=100&&startDate=%s&&endDate=%s"
	# URL for Mode 1
	url_templete_1 = "http://dart.fss.or.kr/dsab002/search.ax?reportName=%s&&maxResults=100&&textCrpNm=%s"
	headers = {'Cookie':'DSAB002_MAXRESULTS=5000;'}
	
	dart_insider_buy_list = []

	#start_day = datetime(2017,11,15)
	#end_day = datetime(2017,12,15)
	start_day = datetime(config_start_year, config_start_month, config_start_day)
	end_day = datetime(config_end_year, config_end_month, config_end_day)
	delta = end_day - start_day

	## 배당
	#report = "%EB%B0%B0%EB%8B%B9"
	# 임원ㆍ주요주주특정증권등소유상황보고서 
	report = "%EC%9E%84%EC%9B%90%E3%86%8D%EC%A3%BC%EC%9A%94%EC%A3%BC%EC%A3%BC%ED%8A%B9%EC%A0%95%EC%A6%9D%EA%B6%8C%EB%93%B1%EC%86%8C%EC%9C%A0%EC%83%81%ED%99%A9%EB%B3%B4%EA%B3%A0%EC%84%9C"

	for i in range(delta.days + 1):

		d = start_day + timedelta(days=i)
		rdate = d.strftime('%Y%m%d')
		print(rdate)
	
		if (config_mode == 0):
			handle = urllib.request.urlopen(url_templete_0 % (report, rdate, rdate))
		# config mode 1
		else:
			handle = urllib.request.urlopen(url_templete_1 % (report, urllib.parse.quote(corp)))
			print("URL" + url_templete_1 % (report, corp))

		data = handle.read()
		soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
		
		table = soup.find('table')
		trs = table.findAll('tr')
		tds = table.findAll('td')
		counts = len(tds)
		print(counts)

		#if counts > 0:
		if counts > 2:
			# Delay operation
			time.sleep(20)
		
			link_list = []
			docid_list = []
			date_list = []
			corp_list = []
			market_list = []
			title_list = []
			reporter_list = []
			tr_cnt = 0
			
			for tr in trs[1:]:
				tr_cnt = tr_cnt + 1
				time.sleep(1)
				if tr_cnt %100 == 0:
					time.sleep(20)
				tds = tr.findAll('td')
				link = 'http://dart.fss.or.kr' + tds[2].a['href']
				date = tds[4].text.strip().replace('.', '-')
				corp_name = tds[1].text.strip()
				market = tds[1].img['title']
				title = " ".join(tds[2].text.split())
				reporter = tds[3].text.strip()
				
				link_list.append(link)
				date_list.append(date)
				corp_list.append(corp_name)
				market_list.append(market)
				title_list.append(title)
				reporter_list.append(reporter)
				#print(corp_name)
				#print(title)
				print(link)

				if ((title == "[기재정정]임원ㆍ주요주주특정증권등소유상황보고서") or (title == "임원ㆍ주요주주특정증권등소유상황보고서")) and (market != "코넥스시장"):

					print(corp_name)
					print(title)
					print(date)
					handle = urllib.request.urlopen(link)
					#print(link)
					data = handle.read()
					soup2 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
					#print(soup2)

					head_lines = soup2.find('head').text.split("\n")
	
					### 보고자에 관한 사항
					re_tree_find1 = re.compile("2.[ ]*보고자에 관한 사항")
					line_num = 0
					line_find = 0
					for head_line in head_lines:
						#print(head_line)
						if (re_tree_find1.search(head_line)):
							line_find = line_num
							break
						line_num = line_num + 1
		
					if(line_find != 0):
			
						line_words = head_lines[line_find+4].split("'")
						#print(line_words)
						rcpNo = line_words[1]
						dcmNo = line_words[3]
						eleId = line_words[5]
						offset = line_words[7]
						length = line_words[9]

						dart = soup2.find_all(string=re.compile('dart.dtd'))
						dart2 = soup2.find_all(string=re.compile('dart2.dtd'))
						dart3 = soup2.find_all(string=re.compile('dart3.xsd'))

						if len(dart3) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart3.xsd"
						elif len(dart2) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart2.dtd"
						elif len(dart) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart.dtd"
						else:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=0&offset=0&length=0&dtd=HTML"  
						
						print(link2)

						handle = urllib.request.urlopen(link2)
						print(handle)
						data = handle.read()
						soup3 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

						tables = soup3.findAll("table")
						trs = tables[0].findAll('tr')
						print("trs : ", len(trs))

						# tr[0]  : 보고구분
						# tr[1]  : 성명
						# tr[2]  : 성명
						# tr[3]  : 주소
						# tr[4]  : 발행회사와의 관계
						# tr[5]  : 발행회사와의 관계
						# tr[6]  : 발행회사와의 관계
						# tr[7]  : 업무상 연락처 및 담당자
						# tr[8]  : 업무상 연락처 및 담당자
						# tr[9]  : 업무상 연락처 및 담당자
						# tr[10] : 업무상 연락처 및 담당자

						# 성명
						tds = trs[1].findAll('td')
						buyer_name = tds[2].text.strip()
						print(buyer_name)

						# 생년월일
						tds = trs[2].findAll('td')
						buyer_date = tds[1].text.strip()
						print(buyer_date)

						# 임원여부
						tds = trs[4].findAll('td')
						buyer_info1 = tds[2].text.strip()
						print(buyer_info1)

						# 직위명
						tds = trs[4].findAll('td')
						buyer_info2 = tds[4].text.strip()
						print(buyer_info2)

						# 선임일
						tds = trs[5].findAll('td')
						buyer_info3 = tds[1].text.strip()
						print(buyer_info3)

						# 퇴임일
						tds = trs[5].findAll('td')
						buyer_info4 = tds[3].text.strip()
						print(buyer_info4)

					### 특정증권등의 소유상황
					re_tree_find1 = re.compile("3.[ ]*특정증권등의 소유상황")
					line_num = 0
					line_find = 0
					for head_line in head_lines:
						#print(head_line)
						if (re_tree_find1.search(head_line)):
							line_find = line_num
							break
						line_num = line_num + 1
		
					if(line_find != 0):
			
						line_words = head_lines[line_find+4].split("'")
						#print(line_words)
						rcpNo = line_words[1]
						dcmNo = line_words[3]
						eleId = line_words[5]
						offset = line_words[7]
						length = line_words[9]

						dart = soup2.find_all(string=re.compile('dart.dtd'))
						dart2 = soup2.find_all(string=re.compile('dart2.dtd'))
						dart3 = soup2.find_all(string=re.compile('dart3.xsd'))

						if len(dart3) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart3.xsd"
						elif len(dart2) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart2.dtd"
						elif len(dart) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart.dtd"
						else:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=0&offset=0&length=0&dtd=HTML"  
						
						print("33333:", link2)

						handle = urllib.request.urlopen(link2)
						print(handle)
						data = handle.read()
						soup3 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

						tables = soup3.findAll("table")
						print("Number of tables : ", len(tables))
						#trs = tables[0].findAll('tr')
						#print("trs : ", len(trs))

						# tables[0] : 소유 특정증권등의 수 및 소유비율 
						# tables[1] : 특정증권등의 종류별 소유내역 
						# tables[2] : 특정증권등의 종류별 소유내역 
						# tables[3] : 세부변동내역
			
						trs = tables[3].findAll('tr')
						print("trs : ", len(trs))
						
						# tr[0:1]   : 보고사유
						# tr[2:-2]   : 
						# tr[-1]  : 합계

						# td[0] : 보고사유
						# td[1] : 변동일
						# td[2] : 특정증권등의종류
						# td[3] : 소유주식수 - 변동전
						# td[4] : 소유주식수 - 증감
						# td[5] : 소유주식수 - 변동후
						# td[6] : 취득처분단가
						# td[7] : 비고

						
						for tr in trs[2:-1]:
						
							dart_insider_buy_sublist = []
							tds = tr.findAll('td')
							print("tds : ", len(tds))
				
							# 보고사유
							buyer_detail1 = tds[0].text.strip()
							print(buyer_detail1)
							# 변동일
							buyer_detail2 = tds[1].text.strip()
							print(buyer_detail2)
							# 특정증권등의 종류
							buyer_detail3 = tds[2].text.strip()
							print(buyer_detail3)
							# 소유주식수 - 변동전
							buyer_detail4 = tds[3].text.strip()
							print(buyer_detail4)
							# 소유주식수 - 증감
							buyer_detail5 = tds[4].text.strip()
							print(buyer_detail5)
							# 소유주식수 - 변동후
							buyer_detail6 = tds[5].text.strip()
							print(buyer_detail6)
							# 취득처분단가
							buyer_detail7 = tds[6].text.strip()
							print(buyer_detail7)
							# 비고
							buyer_detail8 = tds[7].text.strip()
							print(buyer_detail8)

							dart_insider_buy_sublist.append(date)
							dart_insider_buy_sublist.append(corp_name)
							dart_insider_buy_sublist.append(market)
							dart_insider_buy_sublist.append(title)
							dart_insider_buy_sublist.append(link)
							dart_insider_buy_sublist.append(buyer_name)
							dart_insider_buy_sublist.append(buyer_date)
							dart_insider_buy_sublist.append(buyer_info1)
							dart_insider_buy_sublist.append(buyer_info2)
							dart_insider_buy_sublist.append(buyer_info3)
							dart_insider_buy_sublist.append(buyer_info4)
							dart_insider_buy_sublist.append(buyer_detail1)
							dart_insider_buy_sublist.append(buyer_detail2)
							dart_insider_buy_sublist.append(buyer_detail3)
							dart_insider_buy_sublist.append(buyer_detail4)
							dart_insider_buy_sublist.append(buyer_detail5)
							dart_insider_buy_sublist.append(buyer_detail6)
							dart_insider_buy_sublist.append(buyer_detail7)
							dart_insider_buy_sublist.append(buyer_detail8)
							
							dart_insider_buy_list.append(dart_insider_buy_sublist)
				
	cur_dir = os.getcwd()
	
	# Write an Excel file

	#workbook = xlsxwriter.Workbook(workbook_name)
	#if os.path.isfile(os.path.join(cur_dir, workbook_name)):
	#	os.remove(os.path.join(cur_dir, workbook_name))
	workbook = xlsxwriter.Workbook(workbook_name)

	worksheet_result = workbook.add_worksheet('result')
	filter_format = workbook.add_format({'bold':True,
										'fg_color': '#D7E4BC'
										})

	percent_format = workbook.add_format({'num_format': '0.00%'})

	roe_format = workbook.add_format({'bold':True,
									  'underline': True,
									  'num_format': '0.00%'})

	num_format = workbook.add_format({'num_format':'0.00'})
	num2_format = workbook.add_format({'num_format':'#,##0'})
	num3_format = workbook.add_format({'num_format':'#,##0.00',
									  'fg_color':'#FCE4D6'})

	worksheet_result.set_column('A:A', 10)
	worksheet_result.set_column('B:B', 15)
	worksheet_result.set_column('C:C', 15)
	worksheet_result.set_column('D:D', 20)
	worksheet_result.set_column('H:H', 15)
	worksheet_result.set_column('I:I', 15)
	worksheet_result.set_column('J:J', 15)
	worksheet_result.set_column('K:K', 15)


	worksheet_result.write(0, 0, "날짜", filter_format)
	worksheet_result.write(0, 1, "회사명", filter_format)
	worksheet_result.write(0, 2, "분류", filter_format)
	worksheet_result.write(0, 3, "제목", filter_format)
	worksheet_result.write(0, 4, "link", filter_format)
	worksheet_result.write(0, 5, "성명", filter_format)
	worksheet_result.write(0, 6, "생년월일/사업자번호", filter_format)
	worksheet_result.write(0, 7, "임원여부", filter_format)
	worksheet_result.write(0, 8, "직위명", filter_format)
	worksheet_result.write(0, 9, "선임일", filter_format)
	worksheet_result.write(0, 10, "퇴임일", filter_format)
	worksheet_result.write(0, 11, "보고사유", filter_format)
	worksheet_result.write(0, 12, "변동일", filter_format)
	worksheet_result.write(0, 13, "특정증권등의 종류", filter_format)
	worksheet_result.write(0, 14, "변동전", filter_format)
	worksheet_result.write(0, 15, "증감", filter_format)
	worksheet_result.write(0, 16, "변동후", filter_format)
	worksheet_result.write(0, 17, "취득처분단가", filter_format)
	worksheet_result.write(0, 18, "비고", filter_format)

	for k in range(len(dart_insider_buy_list)):
		worksheet_result.write(k+1,0, dart_insider_buy_list[k][0])
		worksheet_result.write(k+1,1, dart_insider_buy_list[k][1])
		worksheet_result.write(k+1,2, dart_insider_buy_list[k][2])
		worksheet_result.write(k+1,3, dart_insider_buy_list[k][3])
		worksheet_result.write(k+1,4, dart_insider_buy_list[k][4])
		worksheet_result.write(k+1,5, dart_insider_buy_list[k][5])
		worksheet_result.write(k+1,6, dart_insider_buy_list[k][6])
		worksheet_result.write(k+1,7, dart_insider_buy_list[k][7])
		worksheet_result.write(k+1,8, dart_insider_buy_list[k][8])
		worksheet_result.write(k+1,9, dart_insider_buy_list[k][9])
		worksheet_result.write(k+1,10, dart_insider_buy_list[k][10])
		worksheet_result.write(k+1,11, dart_insider_buy_list[k][11])
		worksheet_result.write(k+1,12, dart_insider_buy_list[k][12])
		worksheet_result.write(k+1,13, dart_insider_buy_list[k][13])
		worksheet_result.write(k+1,14, dart_insider_buy_list[k][14])
		worksheet_result.write(k+1,15, dart_insider_buy_list[k][15])
		worksheet_result.write(k+1,16, dart_insider_buy_list[k][16])
		worksheet_result.write(k+1,17, dart_insider_buy_list[k][17])
		worksheet_result.write(k+1,18, dart_insider_buy_list[k][18])

	workbook.close()


# Main
if __name__ == "__main__":
	main()


