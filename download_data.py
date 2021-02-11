from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import os
import re
import datetime
import glob



dtypes={'Code':str, 'Name':str, 
	'Open':int, 'High':int, 'Low':int, 'Close':int, 'Volume':int, 'Amount':int,
	'Changes':int, 'ChangeCode':str, 'ChagesRatio':float, 'Marcap':int, 'Stocks':int,
	'MarketId':str, 'Market':str, 'Dept':str,
	'Rank':int}


download_folder = "C:\\update_xls"

if not os.path.exists("C:\\update_xls"):
	os.mkdir("C:\\update_xls")

chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "C:\\update_xls"}
chromeOptions.add_experimental_option("prefs",prefs)
chromedriver = "C:\\Windows\\chromedriver.exe"


driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeOptions)
driver.get("http://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201020101")




def get_daily_data(last_date):
	next_date = last_date + datetime.timedelta(days=1)
	year =  next_date.year
	month = next_date.month
	day = next_date.day

	print("Last day of data: "+last_date.strftime("%Y-%m-%d"))
	print("Crawl       From: "+next_date.strftime("%Y-%m-%d"))

	if datetime.datetime.now().hour > 16:
		end_date  = datetime.datetime.now()
	else:
		end_date = datetime.datetime.now() - datetime.timedelta(days=1)

	print("              To: "+end_date.strftime("%Y-%m-%d"))

	daterange = pd.date_range(next_date, end_date)

	if len(daterange) >= 1:
		
		for d in daterange:
			print(d)
			date_input = driver.find_element_by_name("trdDd")
			print(date_input.get_attribute("value"))
			date_input.send_keys(Keys.CONTROL + "a")
			date_input.send_keys(Keys.DELETE)
			date_input.send_keys(d.strftime("%Y%m%d"))

			time.sleep(1)

			search_button = driver.find_element_by_name("search")
			search_button.send_keys(Keys.ENTER)

			time.sleep(5)
			print("Download Xls")

			b = driver.find_element_by_class_name("CI-MDI-UNIT-DOWNLOAD")
			b.send_keys(Keys.ENTER)

			time.sleep(1)


			filedown_wrap = driver.find_element_by_class_name("filedown_wrap")
			xls_down_div = filedown_wrap.find_element_by_tag_name("div")
			xls_down = xls_down_div.find_element_by_tag_name("a")
			xls_down.send_keys(Keys.ENTER)

			time.sleep(5)

			list_of_files = glob.glob('C:\\update_xls\\*.xlsx')
			latest_file = max(list_of_files, key=os.path.getctime)
			print(latest_file)
			os.rename(latest_file, "C:\\update_xls\\"+d.strftime("%Y-%m-%d.xlsx"))






if __name__ == "__main__":

	year_re = re.compile(r'marcap-(\d\d\d\d)\.csv\.gz')

	files = list()
	for f in os.listdir('data'):
		files.append(f)

	print("There are %d files in /data folder." % len(files))


	last_file = files[-1]
	print(last_file)

	df = pd.read_csv('data/'+last_file)
	print(df.columns)

	last_date = df['Date'].values[-1]
	last_date = pd.to_datetime(last_date)
	next_date = last_date+datetime.timedelta(days=1)

	if (next_date < datetime.datetime.now() and datetime.datetime.now().hour > 16) or (next_date < (datetime.datetime.now() - datetime.timedelta(days=1))):
		print("start crawling...")
		get_daily_data(last_date)
	else:
		print("no need to update.")