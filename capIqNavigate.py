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

# This report type returns the correct html id for each download_type in Report Builder
def getReportType(download_type):
	if(download_type == "customer"):
		download_id = "RepBldrTemplateImg1126682"
	elif(download_type == "supplier"):
		download_id = "RepBldrTemplateImg1126681"
	elif(download_type == "corporate_tree"):
		download_id = "RepBldrTemplateImg1126743"
	else:
		print "************************"
		print "Allowed report types: customer / supplier" 
		exit("Unrecognized report type")

	return download_id

# This starts Firefox with the profile appropriate for Capital IQ download
# The start_page should be an appropriate Capital IQ URL
def capiqInitialize(start_page):
	profile = FirefoxProfile()
        # No SaveToLocation prompt when downloading .xls files
	profile.set_preference("browser.helperApps.neverAsk.saveToDisk",\
		       "application/vnd.ms-excel")
        # Never update firefox
        profile.set_preference("app.update.auto", "false")
        profile.set_preference("app.update.enabled", "false")
        profile.set_preference("app.update.silent", "false")

	driver = webdriver.Firefox(firefox_profile = profile)
	driver.get(start_page)
	sleep(3)
	print "Browser loaded"

	return driver

# This logs the user into Capital IQ
# You must specify the user_id and password
def capiqLogin(driver, user_id, user_password):
	login_success = False
	try:
		username = driver.find_element(by=By.ID, value="username")
		password = driver.find_element(by=By.NAME, value="password")
		signin = driver.find_element(by=By.ID, value="myLoginButton")
		sleep(5)

		username.send_keys(user_id)
		password.send_keys(user_password)
		signin.click()
		print "Login info entered. Signing in."
		sleep(10)

		if "Problem loading page" in driver.title:
			print "Login failed", driver.title

		elif "Report Builder" or "Company Screening" in driver.title:
			print "Login successful: ", driver.title
			login_success = True
		else:
			print "Login failed: ", driver.title
	except (NoSuchElementException, TimeoutException, UnexpectedAlertPresentException):
		print "Login failed: ", driver.title

	return driver, login_success

# For Report Builder, this function returns the number of CIQ IDs which still link to an existing firm in Cap IQ
def getValidFirmCount(driver):
        count = 0
	try:
		count_string = WebDriverWait(driver,15).until(\
			        EC.presence_of_element_located((\
			        By.PARTIAL_LINK_TEXT,\
			        "Companies("))\
                                ).text
		count = int(filter(type(count_string).isdigit, count_string))
			   

	except (TimeoutException, NoSuchElementException):
                pass

	print "%s valid CQ IDs" % (count)

	return int(count)

# This function adds CIQ IDs into the Report Builder dialogue
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

# While in the Report Generation dialog, 
# This function attempts to get the download URL, it returns the download success and the filename
# Otherwise it returns a False status which means still generating
# Or a "Failed" status which means the report generation failed
def downloadFile(driver, batch_no):
	download_success = False
	filename = ""

	try:
		# Switch to the report generation window
		# Ensure that this window is in focus
		for handle in driver.window_handles:
			driver.switch_to.window(handle)
			if driver.title[:12] ==  "Capital IQ Report Status":
				break

		# Get file-name of the file in the first row of the report generation window
                sleep(5)
		filename_element = WebDriverWait(driver,5).until(\
                    	  	   EC.presence_of_element_located((\
                           	   By.XPATH, "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[1]/div[1]")))
		filename = filename_element.text + ".xls"

		# Get the file-link of the file in the first row of the report generation window
		file_url = WebDriverWait(driver,5).until(\
                   	        EC.presence_of_element_located((\
 		    	        By.XPATH,\
                                "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[3]/span/a"))\
                                ).get_attribute("href")
	
		# If file_url is not empty, then start downloading
		if file_url is not None:
			print ""
			print "Getting %s from url %s" % (filename, file_url)
			driver.get(file_url)
			# print "Downloading batch file #" + str(batch_no)
			download_success = True

	except TimeoutException:
		file_status = WebDriverWait(driver,3).until(\
                    	      EC.presence_of_element_located((\
                              By.XPATH,\
			      "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[1]/td[3]/span"))
			      ).text

                #  This is when report generation fails
		if file_status == "Failed":
			download_success = "Failed"
			print ""

		# print "File status: %s" % (file_status)
		

	return download_success, filename
	

# This function begins the report generation process, calls the download function repeatedly
# Each time it tests if the  report is 1) Successfully generated 2) Still generating 3) Failed
# It returns the download_success and the filename if available
def generateReport(driver, batch_no, min_wait_time, max_wait_time, download_id):
	# Generate Report
	sleep(2)
	filename = ""
	success = False
	generate_report = WebDriverWait(driver,60).until(\
                    	  EC.element_to_be_clickable((\
			  By.ID, download_id))
			  )

	generate_report.click()
	print "Generating Report"


	# Each time, allow for the min download time to elapse 
	# If status == "Failed" or max_wait_time is exceeded, exit loop
        # Return generation failure
	total_wait_time = 0
	print "Seconds waited:",
	while total_wait_time < max_wait_time and success == False: 
		sleep(min_wait_time)
		total_wait_time += min_wait_time
		print str(total_wait_time),
		success, filename = downloadFile(driver, batch_no)

	for handle in driver.window_handles:
		driver.switch_to.window(handle)
		if driver.title[:12] ==  "Capital IQ R":
			driver.close()


	return success, filename 

# This protocol logs the user out of Capital IQ and closes the main window
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

