# Automated downloading of customer-supplier relations from Capital IQ
from sys import argv, exit
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import UnexpectedAlertPresentException 
from selenium.common.exceptions import NoAlertPresentException 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep, time, localtime, strftime
from os import chdir, remove, rename, listdir
from shutil import copy
from math import ceil
from capIqNavigate import capiqInitialize, capiqLogin, capiqLogout, generateReport, downloadFile
from capIqLibrary import createDummyFile, getDownloadName,isDownloadDirClear, moveAllExcelFiles, moveAllPartialFiles, readDownloadDir

def switchToOldScreening(driver):
	try:
		old_switch = WebDriverWait(driver, 30).until(
			     EC.presence_of_element_located((By.ID,"returnToOriginalLinkNonBeta"))
			     )
		old_switch.click()
		sleep(3)
	except (TimeoutException, NoSuchElementException):
		exit("!Exception at switching to old interface")

	return 

def getScreenId(browser_url):

	begin = browser_url.find("UniqueScreenId=")
	end = browser_url.find("&", begin)
	screen_url = browser_url[begin:end]
	screen_id = filter(lambda x: x.isdigit(), screen_url)
	print "Screen ID is %s" % (screen_id)

	return screen_id

def setGicFilter(driver, gic_code):
	try:
		screening_search = WebDriverWait(driver, 30).until(
				   EC.presence_of_element_located((By.ID, "SearchDataPointsAutoCompleteTextBoxPhase2"))
				   )

		screening_search.send_keys(gic_code)
		sleep(1)
		# screening_search.send_keys(Keys.ENTER)
	except (TimeoutException, NoSuchElementException):
		exit("!Exception when entering GIC Code")

	try:
		sleep(5)
		sub_search = WebDriverWait(driver,10).until(
			     EC.presence_of_element_located((By.XPATH,\
                             "/html/body/table/tbody/tr[2]/td[4]/div/form/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td/div/span/div/div[1]/div[2]/a"))
			     )
		sub_search.click()
		sleep(3)
		screening_search.send_keys(Keys.ENTER)

		print "Criterion added"
		sleep(3)
		driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
	except (TimeoutException, NoSuchElementException):
		exit("!Exception at adding GIC criterion")

	try:
		driver.switch_to_frame("CriterionResultsFrame")
		firm_count_elem = WebDriverWait(driver,15).until(
				  EC.element_to_be_clickable((By.ID,
			          "_viewTopControl__numberOfResults"))
				  )
		firm_count = int(filter(lambda x: x.isdigit(), firm_count_elem.text))
	except (TimeoutException, NoSuchElementException):
		exit("!Exception at getting firm count")

	try:
		view_results = WebDriverWait(driver,15).until(
				EC.element_to_be_clickable((By.ID,
				"_viewTopControl__resultsLink"))
				)
		view_results.click()
	except (TimeoutException, NoSuchElementException):
		exit("!Exception at viewing results")

	return firm_count

def setTemplate(driver, option):
	sleep(3)
	template_name = "Invalid"
	if "Screening Results" not in driver.title:
		exit("!Exception. Not in screening page")

	drop_down = WebDriverWait(driver,15).until(
		    EC.element_to_be_clickable((By.ID,
		    "_displayOptions_Displaysection1_SelectedTemplate"))
		    )
	drop_down.click()
	sleep(3)

	if option == 1:
		drop_down.send_keys("mass_extraction")
		template_name = "mass_extraction_1"

	set_template=WebDriverWait(driver,15).until(
	     	     EC.element_to_be_clickable((By.ID,
	   	     "_displayOptions_Displaysection1_GoButton"))
		     )
	set_template.click()


	return template_name

def changePageNo(driver, page_no):
	pass

	return

def renameBatchFile(batch_no, download_path, download_name, company_names_info):
	final_name = "Invalid"
	rename_success = False
	wait_time = 10
	total_wait_time = 0

	# Check if the download is compelete by checking for .part files
	download_success = False

	while download_success is False and total_wait_time < 120:
		downloaded_files = listdir(download_path)
		download_success = True
		for downloaded_file in downloaded_files:
			if downloaded_file[-5:] == ".part":
				print "Download incomplete. Total time waited: ", str(total_wait_time), "seconds"
				download_success = False
				sleep(wait_time)
				total_wait_time += wait_time
				break

	# If download fails, return "Invalid" and rename_success = False
	if download_success is False:
		return final_name, rename_success

	# When download is complete, search for the downloaded file to rename it
	for index in range(len(downloaded_files)):
		actual_name = str(downloaded_files[index])

		## Use 2 methods to generate the filename, if they agree, rename the file
		## if not use batch no to rename the file, but mark the file with "star_"

		# method 1: matching firm name with master table
		# true_name is what the file will be renamed to
		if actual_name == download_name:
			true_name = getTrueName(actual_name, company_names_info)

		# method 2: check download batch no and file name
			if actual_name[-13:] == "Suppliers.xls":
				batch_filename = "suppliers_batch_" + str(batch_no) + ".xls"
			elif actual_name[-13:] == "Customers.xls":		
				batch_filename = "customers_batch_" + str(batch_no) + ".xls"
			elif actual_name[-17:] == "CorporateTree.xls":
				batch_filename = "corporateT_batch_" + str(batch_no) + ".xls"

			try:
				# Where the 2 methods agree or getTrueName returns "Invalid"
				if true_name == batch_filename or true_name is "Invalid":
					rename(actual_name, batch_filename)
					final_name = batch_filename

					print "Batch agrees with master, File renamed to %s"\
					      % (batch_filename)

				# Where the 2 methods disagree
				else:
					rename(actual_name, "star_" + batch_filename)
					final_name = "star_" + batch_filename
					print "Batch different from master, File renamed to %s"\
					      % ("star_" + batch_filename)

				rename_success = True
				break
			except WindowsError:
				print batch_filename + " used by another file"
				rename_success = True
				break
	if rename_success == False:
		print "Rename process unsuccessful"

	return final_name, rename_success

# 0: Program starts: check arguments
print "********** Capital IQ mass screening downloader *********"
if (len(argv) < 2):
	print "Format: python mass_screening.py <GIC_CODE>"
	exit("Insufficient arguments")

# 1: Check if download directory is free of excel files
download_path = readDownloadDir("C:/Selenium/capitaliq/download_dir.txt")
if download_path == "Invalid":
	exit("Invalid download directory")

if isDownloadDirClear(download_path) is False:
	exit("Download dir is not clear. Remove all .xls and .xlsx files")


# 2: Initialize the brower and load Capital IQ 
login_attempts = 0
login_success = False
while login_attempts < 3 and login_success is False:
	login_attempts += 1	
	screening_page = "https://www.capitaliq.com/ciqdotnet/Screening/ScreenBuilder.aspx?clear=all&returnToOriginal=1"
	driver = capiqInitialize(screening_page)
	main_window = driver.current_window_handle
	print "Capital IQ website loaded"
	print "Login attempt #%d" % (login_attempts)
	driver, login_success = capiqLogin(driver, "davinchor@nus.edu.sg", "GPNm0nster")

	if login_success is False:
		print "Close browser. Wait one minute"
		sleep(60)
		driver.quit()

if login_attempts == 3:
	exit("Login attempts limit exceeded.")

wait_time = 0
while True:
	if "Company Screening" in driver.title:
		print "Company Screening loaded"
		break
	elif wait_time > 60:
		exit("Timeout on loading Company Screening page")
	else:
		sleep(5)
		wait_time += 5

# 3. Set filter to target GIC code and get number of firms
target_gic = argv[1]
total_firm_count = setGicFilter(driver, target_gic)
print "Filter set: %d firms in %s" % (total_firm_count, target_gic)

# 4. Set correct variable template
template_name = setTemplate(driver, 1)
if template_name == "Invalid":
	exit("Invalid template name")
print "Template set: %s" % (template_name)

# 5. Create download list, one file for every 10,000 firms
total_files = int(ceil(float(total_firm_count)/10000.0))
download_list = range(1, total_files+1)
print "Download list: %s" % (str(download_list))

for download_no in download_list:
	if download_no != 1:
		changePageNo(download_no)	

	success = False
	attempt_no = 0
	while success == False and attempt_no < 3:
		success, filename = generateReport(driver, 0, 30, "_displayOptions_Displaysection1_ReportingOptions_GoButton")
		attempt_no += 1

print "Script End"
