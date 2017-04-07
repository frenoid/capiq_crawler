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
from capIqLibrary import createDummyFile, getDownloadName,isDownloadDirClear, moveAllExcelFiles, moveAllPartialFiles, readDownloadDir, checkMakeDir, checkDownloadComplete

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
		screen_id = getScreenId(driver.current_url)
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

	return firm_count, screen_id

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

def changePageNo(driver, page_no, screen_id):

	# Switch to correct page of 50,000 firms
	# For firm count > 50000
	if (page_no-1)%5==0 and page_no != 1:

		# Return to the screen selection page
		assert "Screening Results" in driver.title
		screen_url = "https://www.capitaliq.com/CIQDotNet/Screening/ScreenBuilder.aspx?UniqueScreenId=" + screen_id + "&clear=all&returntooriginal=1#"
		driver.get(screen_url)

		# Switch to the correct set of 50,000 firms
		sleep(10)
		driver.switch_to_frame("CriterionResultsFrame")
		view_range_menu = WebDriverWait(driver, 15).until(
				  EC.element_to_be_clickable((By.ID,
				  "_viewTopControl__range"))
				  )
		if page_no == 6:
			view_range_menu.send_keys("50001")
		elif page_no == 11:
			view_range_menu.send_keys("100001")
		elif page_no == 16:
			view_range_menu.send_keys("150001")
		elif page_no == 21:
			view_range_menu.send_keys("200001")
		elif page_no == 26:
			view_range_menu.send_keys("250001")
		elif page_no == 31:
			view_range_menu.send_keys("300001")
		elif page_no == 36:
			view_range_menu.send_keys("350001")
		elif page_no == 41:
			view_range_menu.send_keys("400001")
		elif page_no == 46:
			view_range_menu.send_keys("450001")
		# View results
		view_results = WebDriverWait(driver,15).until(
			       EC.element_to_be_clickable((By.ID,
			       "_viewTopControl__resultsLink"))
			       )
		view_results.click()

	# Switch to correct set of 0 - 50000 firms
	try:
		export_menu = WebDriverWait(driver, 30).until(
			      EC.element_to_be_clickable((By.ID,
			      "_displayOptions_Displaysection1_ReportingOptions_NumberOfTargetsExcel"))
			      )
		if (page_no % 5) != 1:
			export_menu.click()
			if page_no % 5 == 2:
				export_menu.send_keys("10001")
			elif page_no % 5 == 3:
				export_menu.send_keys("20001")
			elif page_no % 5 == 4:
				export_menu.send_keys("30001")
			elif page_no % 5 == 0:
				export_menu.send_keys("40001")
			export_menu.send_keys(Keys.ENTER)
	finally:
		pass

	
	print "Switched to %d th firms" % ((page_no-1)*10000)


	return

def renameMassFile(download_path, download_name, gic_code, page_no, page_total):
	rename_success = False
	final_name = gic_code + "_" + str(page_no) + "_of_"\
		     + str(page_total) + ".xls"
	entries = listdir(download_path)

	for entry in entries:
		if entry == download_name:
			rename(download_path+"/"+entry,\
			       download_path+"/"+final_name)
			print "%s renamed to %s" % (entry, final_name)
			rename_success = True
	
	return rename_success, final_name

# 0: Program starts: check arguments
print "********** Capital IQ mass screening downloader *********"
if (len(argv) < 2):
	print "Format: python mass_screening.py <GIC_CODE>"
	exit("Insufficient arguments")

target_gic = argv[1]

# 1: Check if download directory is free of excel files
download_path = readDownloadDir("C:/Selenium/capitaliq/download_dir.txt")
if download_path == "Invalid":
	exit("Invalid download directory")
if isDownloadDirClear(download_path) is False:
	exit("Download dir is not clear. Remove all .xls and .xlsx files")

final_path = download_path + "/" + target_gic
checkMakeDir(final_path)


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
total_firm_count, screen_id = setGicFilter(driver, target_gic)
print "Filter set: %d firms in %s" % (total_firm_count, target_gic)

# 4. Set correct variable template
template_name = setTemplate(driver, 1)
if template_name == "Invalid":
	exit("Invalid template name")
print "Template set: %s" % (template_name)

# 5. Create download list, one file for every 10,000 firms
total_files = int(ceil(float(total_firm_count)/10000.0))
download_list = range(1, total_files+1)
failed_page_downloads = []
print "Download list: %s" % (str(download_list))
print "***** Download Commenced *****"

for download_no in download_list:
	print "=== File %d of %d ===" % (download_no, len(download_list))
	try:
		# Change to approriate page number
		if download_no != 1:
			changePageNo(driver, download_no, screen_id)	

		# Get download link and download
		success = False
		attempt_no = 0
		while success == False and attempt_no < 3:
			success, filename = generateReport(driver, 0, 30, "_displayOptions_Displaysection1_ReportingOptions_GoButton")
			attempt_no += 1

		# Ensure the download is done	
		sleep(5)
		download_complete = False
		total_wait_time = 0
		while download_complete == False and total_wait_time < 120:
			download_complete = checkDownloadComplete(download_path)
			if download_complete == False:
				sleep(10)
				print "Download incomplete. Time elapsed %d"\
				      % (total_wait_time)
				total_wait_time += 10

		# Rename to download file accordingly
		renameMassFile(download_path, filename,\
			       target_gic,download_no, len(download_list))

	except(TimeoutException, NoSuchElementException, UnexpectedAlertPresentException) as exception_type:
		print "!Exception of type", exception_type, "encountered"
		failed_page_downloads.append(download_no)
		driver.switch_to_window(main_window)
		driver.get(current_url)
		sleep(10)
				
	finally:
		excel_files_moved = moveAllExcelFiles(download_path, final_path)
		print "%d .xls files moved" % (excel_files_moved)
		partial_files_moved = moveAllPartialFiles(download_path, final_path)
		print "%d .part files moved" % (partial_files_moved)
		driver.switch_to_window(main_window)

print "Failed pages:", str(failed_page_downloads)
capiqLogout(driver, main_window)
print "Script End"
