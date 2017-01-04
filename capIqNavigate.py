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
import pyperclip
from time import sleep

def getReportType(download_type):
	if(download_type == "customer"):
		download_id = "RepBldrTemplateImg1126682"
	elif(download_type == "supplier"):
		download_id = "RepBldrTemplateImg1126681"
	else:
		print "************************"
		print "Allowed report types: customer / supplier" 
		exit("Unrecognized report type")

	return download_id


def capiqInitialize(report_page):
	profile = FirefoxProfile()
	profile.set_preference("browser.helperApps.neverAsk.saveToDisk",\
		       "application/vnd.ms-excel")

	driver = webdriver.Firefox(firefox_profile = profile)
	driver.get(report_page)
	print "Browser loaded"

	return driver

def capiqLogin(driver, user_id, user_password):
	login_success = False
	try:
		username = driver.find_element(by=By.ID, value="username")
		password = driver.find_element(By.NAME, "password")
		signin = driver.find_element(by=By.ID, value="myLoginButton")

		username.send_keys(user_id)
		password.send_keys(user_password)
		signin.click()
		print "Login info entered"

		WebDriverWait(driver,15).until(EC.title_contains("Report Builder"))
		print "Login successful: " + driver.title
		login_success = True

	except (NoSuchElementException, TimeoutException, UnexpectedAlertPresentException):
		print "Login failed. Close browser."
		driver.close()

	return login_success

def getValidFirmCount(driver):
	try:
		count_element = WebDriverWait(driver,15).until(\
			        EC.presence_of_element_located((\
			        By.PARTIAL_LINK_TEXT,\
			        "Companies(")))

		count_string = count_element.text
		count = filter(type(count_string).isdigit, count_string) 
			   
		print "%s valid CQ IDs" % (count)

	except (TimeoutException, NoSuchElementException):
		count = 0
		print "0 valid CQ IDs"

	return int(count)

def addFirms(driver, batch_list):
	add_firm = WebDriverWait(driver,30).until(\
		   EC.presence_of_element_located((\
		   By.ID,"_rptOpts__rptOptsDS__optsDs__optsTog__esLink")))
	print "Report Builder loaded"
	add_firm.click()

	# Enter IDs into the Search box and search
	search_box = WebDriverWait(driver,15).until(\
		     EC.presence_of_element_located((\
		     By.CLASS_NAME, "es-searchinput")))

	search_string = "\n".join(batch_list)
	pyperclip.copy(search_string)

	search_box.click()
	ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
	sleep(1)	

	search_submit = driver.find_element(By.CLASS_NAME,"entitysearch-search")
	sleep(3)

	search_submit.click()

	# Get the number of valid firms and add to report
	sleep(5)
	valid_firm_count = getValidFirmCount(driver)

	
		
	add_to_report = WebDriverWait(driver,15).until(\
			EC.presence_of_element_located((\
			By.ID,\
			"_rptOpts__rptOptsDS__optsDs__optsTog_float_esModal__esSaveCancel__saveBtn")))
	add_to_report.click()
	print "Firms added"

	return valid_firm_count

def downloadFile(driver, batch_no):
	download_success = False
	filename = ""

	try:
		# Get file-name of download file
		filename_element = WebDriverWait(driver,3).until(\
                    	  	   EC.presence_of_element_located((\
                           	   By.XPATH, "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[1]/div[1]")))
		filename = filename_element.text + ".xls"

		file_link = WebDriverWait(driver,3).until(\
                   	    EC.presence_of_element_located((\
 		    	    By.XPATH, "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[3]/span/a")))
	
		# Check if URL is empty, if not download	    
		file_url = file_link.get_attribute("href")
		if file_url is not None:
			print "Getting %s from url %s" % (filename, file_url)
			driver.get(file_url)
			print "Downloading batch file #" + str(batch_no)
			download_success = True
		else:
			download_success = False
			

	except TimeoutException:
		file_status = WebDriverWait(driver,1).until(\
                    	      EC.presence_of_element_located((\
                              By.XPATH, "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[3]/span")))
		print "File status: %s" % (file_status.text)
		
		if file_status.text == "Failed":
			download_success = "Failed"
		else:
			download_success = False
	

	return download_success, filename
	

def generateReport(driver, batch_no, min_wait_time, download_id):
	# Generate Report
	sleep(2)
	filename, success = "", False
	generate_report = WebDriverWait(driver,15).until(\
                    	  EC.presence_of_element_located((\
			  By.ID, download_id)))

	generate_report.click()
	print "Generating Report"

	# Switch to the Download progress windows
	for handle in driver.window_handles:
		driver.switch_to.window(handle)
		if driver.title[:12] ==  "Capital IQ R":
			break

	# 3 checks for the genreation completeness 
	# Each time, allow for the min download time to elapse 
	# If status == "Failed", exit loop, return generation failure
	download_attempts = 0
	while download_attempts < 3 and success != True and success != "Failed": 
		download_attempts += 1
		print "Download attempt #", str(download_attempts)
		sleep(min_wait_time)
		success, filename = downloadFile(driver, batch_no)

	for handle in driver.window_handles:
		driver.switch_to.window(handle)
		if driver.title[:12] ==  "Capital IQ R":
			driver.close()

	if success is "Failed":
		success = False

	return success, filename 

def capiqLogout(driver, main_window):
	try:
		driver.switch_to.window(main_window)
		logout_link = WebDriverWait(driver,15).until(\
		              EC.presence_of_element_located((By.LINK_TEXT,"Logout")))
        	logout_link.click()
		sleep(2)
		print "Logging out" 

	except (TimeoutException, UnexpectedAlertPresentException, NoSuchElementException):
		print "!Exception encountered during logout"

	finally:
		driver.close()
		print "Exiting browser"

	return

