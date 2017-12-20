# DART_dividends

Get dividend information from postings in DART

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

Description of python codes (in Korean)

https://blog.naver.com/jaden-agent/221165184133

https://blog.naver.com/jaden-agent/221166737850 
 
