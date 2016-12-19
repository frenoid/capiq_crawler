# Automated downloading of customer-supplier relations from Capital IQ
from sys import argv, exit
from openpyxl import load_workbook
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
from capIqLibrary import getCompanyNamesInfo, getTrueName
from time import sleep, time, localtime, strftime
from os import chdir, remove, rename, listdir
from shutil import copy
from math import ceil
import pyperclip

def getReportType(download_type):
	print download_type
	if(download_type == "customer"):
		download_id = "RepBldrTemplateImg1126682"
	elif(download_type == "supplier"):
		download_id = "RepBldrTemplateImg1126681"
	else:
		print "************************"
		print "Allowed report types: customer / supplier" 
		exit("Unrecognized report type")

	return download_id

def getFirmList(ids_file):
	
	# Read the file of firm IDs
	company_names_info = {}
	ids_file = argv[1]

	# Reading an excel file
	if ids_file[-4:] == "xlsx":
		company_names_info = getCompanyNamesInfo(ids_file)

	elif ids_file[-4:] == ".txt" or ids_file[-4:] == ".xls":
		exit(".txt and .xls are no longer supported")

	else:
		print "%s is an unknown file format" % (ids_file)
		exit()


	return company_names_info


def getDownloadList(firm_list, list_of_args):
	# Calculate the number of batches to download 
	if argv[3] > 0:
		batch_size = int(argv[3])
	else:
		exit("Invalid batch size")

	batch_total = int(ceil(float(len(firm_list))/batch_size))
	print str(len(firm_list)) + " Firms and Batch size of "\
	      + str(batch_size) + " produces " + str(batch_total) + " batches"
	print "***"

	print "Step 2: Selecting firms to download"

	# Arg4 can be either "all", "list", or an integer
	download_list = []

	# Get all batches 
	if(argv[4] == "all"):
		download_list = range(1, batch_total+1)
	
	# Get a list of batches
	elif(argv[4] == "list"):
		download_list.extend(argv[5:]) 
		for download_batch in download_list:
			download_batch = int(download_batch)
			if(download_batch > batch_total):
				print "Batch #%d exceeds total %d"\
				      % (download_batch, batch_total)
				exit("Batch exceeds batch range")

	# Get a range of batches if arg4 is an integer
	elif(argv[4] > 0):
		batch_start = int(argv[4])

		# arg 5 permits an integer or "end"
		if(argv[5] == "end"):
			batch_last = batch_total
		else:
			batch_last = int(argv[5])
		
		# Sanity check before generating download list
		if(batch_start <= batch_last):
			download_list.extend(range(batch_start, batch_last+1))
		elif(batch_last > batch_total):
			exit("Last batch exceeds available batch range")

	# If arg4 doesn't not match any cases	
	else:
		exit("Invalid start batch argument")

	print "Report: *%s*, Batch size: %d" % (argv[2], batch_size)
	print download_list
	batch_download_total = len(download_list) 
	print "Preparing to download %d batches" % (batch_download_total)
	print "***"

	return download_list

def capiqInitialize(report_page):
	profile = FirefoxProfile()
	profile.set_preference("browser.helperApps.neverAsk.saveToDisk",\
		       "application/vnd.ms-excel")

	driver = webdriver.Firefox(firefox_profile = profile)
	driver.get(report_page)
	print "Browser loaded"

	return driver


def capiqLogin(driver, user_id, user_password):
	username = driver.find_element(by=By.ID, value="username")
	password = driver.find_element(By.NAME, "password")
	signin = driver.find_element(by=By.ID, value="myLoginButton")

	username.send_keys(user_id)
	password.send_keys(user_password)
	signin.click()
	print "Login info entered"

	WebDriverWait(driver,15).until(EC.title_contains("Report Builder"))
	print "Login successful: " + driver.title

	return driver

def getBatchList(firms_list, batch_size, batch_no):
	
	# Batch creation
	batch_list = []
	batch_start = (batch_no - 1) * batch_size
	batch_end = batch_no * batch_size
	
	# Determine if last batch, if so then replace batch end number
	# with the last few batches numbers
	if batch_end > len(firm_list):
		batch_end = batch_start + (len(firm_list)%batch_size)

	print "Creating batch list, adding firm #%d to firm #%d"\
	      % (batch_start+1, batch_end)
	batch_list.extend(firm_list[batch_start:batch_end])

	return batch_list

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


def createDummyFile(batch_no, report_type):
	if (report_type == "customer"):
		dummy_file_name = "customers_batch_" + str(batch_no) + ".xls"
	elif (report_type == "supplier"):
		dummy_file_name = "suppliers_batch_" + str(batch_no) + ".xls"
	else:
		dummy_file_name = "unknown_batch_" + str(batch_no) + ".xls"

	copy("C:/Selenium/capitaliq/example_dummy_file.xls",\
	     "C:/Users/faslxkn\Downloads/" + dummy_file_name)

	return dummy_file_name

def generateReport(driver, batch_no, min_wait_time):
	# Generate Report
	success = False
	sleep(2)
	generate_report = driver.find_element(by=By.ID,\
			  value=download_id)
	generate_report.click()
	print "Generating Report"

	# Switch to the Download progress windows
	for handle in driver.window_handles:
		driver.switch_to.window(handle)
		if driver.title[:12] ==  "Capital IQ R":
			break

	# Wait for at least 30s + min_wait_time for report generation to complete
	# If still generating, wait an additional 45 secs
	# If failed, return failed status
	# If link found, download the file

	sleep(min_wait_time)

	# First 30 secs
	try:
		file_link = WebDriverWait(driver,30).until(\
                    	    EC.presence_of_element_located((\
 		            By.LINK_TEXT, "Download")))
			    
		file_url = file_link.get_attribute("href")
		print "Link found! Getting URL " + file_url
		driver.get(file_url)
		print "Downloading batch file #" + str(batch_no)

		success = True

	# 30 secs exceeded
	except TimeoutException:

		# Check for failure
		link = 	WebDriverWait(driver,5).until(\
                    	EC.presence_of_element_located((\
 		        By.XPATH, "/html/body/div[2]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr/td[3]/span")))
		if link.text == "Failed":
			success = False
		
		# Wait an additional 45s
		else:
			print "Long wait time. Wait an additional 30 sec"
			file_link = WebDriverWait(driver,45).until(\
                    	            EC.presence_of_element_located((\
 		                    By.LINK_TEXT, "Download")))	    
			file_url = file_link.get_attribute("href")
			print "Link found! Getting URL " + file_url
			driver.get(file_url)
			print "Downloading batch file #" + str(batch_no)
			success = True

	if success is False:
		print "Report generation failed."

	return success 


def getDownloadName(report_type, valid_firm_count):
	# Produce an expected filename 
	if report_type == "customer":
		download_name = str(valid_firm_count) + "Companies_CompanyCustomers.xls"
	elif report_type == "supplier":
		download_name = str(valid_firm_count) + "Companies_CompanySuppliers.xls"
	else:
	        download_name = str(valid_firm_count) + "Companies.xls"

	return download_name


def renameBatchFile(download_files, download_name, company_names_info):
	rename_success = False

	for index in range(len(downloaded_files)):
		actual_name = str(downloaded_files[index])


		## First method of renaming the file
		# Test the file contents against company_names_info
		if actual_name == download_name:
			true_name = getTrueName(actual_name, company_names_info)
			if true_name is not "Invalid":
				rename(actual_name, true_name)
				print "Using firm name, File renamed to %s" % (true_name)
				rename_success = True
				break

		## Second method of renaming the file
		# Numerical batch numbering
		# For suppliers info	
		if actual_name == download_name and\
		   actual_name[-13:] == "Suppliers.xls":
			try:
				batch_filename = "suppliers_batch_"+\
						 str(batch_no)+".xls"
				rename(str(downloaded_files[index]),\
			       	       batch_filename)
				rename_success = True
				print "Using batch number, File renamed to %s" % (batch_filename)
				break
			except WindowsError:
				print batch_filename + " used by another file"
				rename_success = True
				break

		# For customer info
		elif actual_name == download_name and\
		     actual_name[-13:] == "Customers.xls":		
			try:
				batch_filename = "customers_batch_"+\
						 str(batch_no)+".xls"
				rename(str(downloaded_files[index]),\
			       	batch_filename)
				print "Using batch number, File renamed to %s" % (batch_filename)
				rename_success = True
				break
			except WindowsError:
				print batch_filename + " used by another file"
				rename_success = True
				break
	if rename_success == False:
		print "Rename process unsuccessful"

	return rename_success


def capiqLogout(driver, main_window):
	try:
		driver.switch_to.window(main_window)
		logout_link = WebDriverWait(driver,15).until(\
		      EC.presence_of_element_located((By.LINK_TEXT,"Logout")))
        	logout_link.click()
		sleep(2)
		print "Logging out and exiting"
		driver.close()

	except TimeoutException or UnexpectedAlertPresentException\
	       or NoSuchElementException:
		print "!Exception encountered during logout"

	return

""" Main() """
# 1: Check if there are enough arguments
if (len(argv) < 5):
	print "%d arguments. Minimum is 4" % (len(argv)-1)
	print "Arguments: <file_of_ids> <{customer/supplier}] " +\
	      "<batch_size> <start_batch> [end_batch]"
	exit()

# 2: Set report type
report_type = argv[2]
download_id = getReportType(report_type)

# 3: Read the file of firm IDs
company_names_info = getFirmList(argv[1]) 
firm_list = []
for company in company_names_info:
	firm_list.append(company_names_info[company][0])
firm_list.sort()

batch_size  = int(argv[3])
download_list = getDownloadList(firm_list, argv)
batch_total = int(ceil(float(len(firm_list))/batch_size))

# 4: Initialize the brower and load Capital IQ 
report_page = "https://www.capitaliq.com/ciqdotnet/ReportsBuilder/CompanyReports.aspx"
driver = capiqInitialize(report_page)
main_window = driver.current_window_handle
print "Capital IQ website loaded"

# 5: Login
driver = capiqLogin(driver, "davinchor@nus.edu.sg", "GPNm0nster")

# Initialize batch generation
batch_list = []
batch_processed_count = 0
batch_failed_count = 0
start_time = time()
failed_batches = {}
consec_failure_count = 0
firms_processed_count = 0

# Central loop: Terminates when the last batch is processed 
while batch_processed_count < len(download_list):
	try:
		batch_no = int(download_list[batch_processed_count])
		print "+++++++++++++++++++++++++++++++++++++"
		print "Initialize batch #" + str(batch_no)
		batch_list = getBatchList(firm_list, batch_size, batch_no)

		print "Downloading batch #%d. To download %d of %d"\
		      % (batch_no, len(download_list), batch_total)

		# Break every 15 batches
		if batch_processed_count % 15 == 0 and batch_processed_count > 0:
			print "10 sec break"
			sleep(10)

		# Check if there have been more than 2 or more failures
		# Wait the equivalent number of minutes
		if consec_failure_count >= 2:
			print str(consec_failure_count),\
			      "consecutive failed downloads. Wait 1 minute"
			sleep(60)

		# If not first batch, then Refresh the RB
		if batch_processed_count > 0:
			driver.get(report_page)	
			print "Refresh Report Builder"


		# Add CQ IDs to the Report Generator
		valid_firm_count = addFirms(driver, batch_list)

		# Where there are no valid CQ IDs, create a dummy "No data" .xls file
		# Then proceed to next batch
		if valid_firm_count == 0:
			dummy_file_name = createDummyFile(batch_no, report_type)
			print "Dummy %s was created" % (dummy_file_name)
			print "Next batch"
			continue
	

		# Generate the report
		min_wait_time = (batch_size/5.0)
		generateSuccess = generateReport(driver, batch_no, min_wait_time)

		if generateSuccess is True:
			if consec_failure_count > 0:
				print "Consecutive failures reset to 0"
				consec_failure_count = 0

			
			# Rename the downloaded file, 3 tries allowed 
			download_name = getDownloadName(report_type, valid_firm_count)
			rename_tries = 0
			rename_success = False
			while rename_tries < 2 and rename_success is not True:
				sleep(15)
				rename_tries += 1
				chdir("C:/Users/faslxkn/Downloads")
				downloaded_files = listdir("C:/Users/faslxkn/Downloads")
				rename_success = renameBatchFile(downloaded_files, download_name, company_names_info)

		else:
			print "Batch generation failed"
			consec_failure_count += 1
			print "Consecutive failed attempts: %d" % consec_failure_count
			batch_failed_count += 1
			failed_batches[batch_no] = batch_list
			continue

		# Ensure focus is on main window
		driver.switch_to.window(main_window)

	# Exception handling
	except TimeoutException:
		print "!Exception: Timeout, proceed to next batch"
		consec_failure_count += 1
		print "Consecutive failed attempts: %d" % consec_failure_count
		batch_failed_count += 1
		failed_batches[batch_no] = batch_list
		continue
	except UnexpectedAlertPresentException:
		print "!Exception: Unexpected alert, proceed to next batch"
		consec_failure_count += 1
		print "Consecutive failed attempts: %d" % consec_failure_count
		batch_failed_count += 1
		failed_batches[batch_no] = batch_list
		continue
	except NoSuchElementException:
		print "!Exception: Element not found, proceed to next batch"
		consec_failure_count += 1
		print "Consecutive failed attempts: %d" % consec_failure_count
		batch_failed_count += 1
		failed_batches[batch_no] = batch_list
		continue
	except WebDriverException:
		print "!Exception: WebDriverException, proceed to next batch"
		consec_failure_count += 1
		print "Consective failed attempts: %d" % consec_failure_count
		batch_failed_count += 1
		failed_batches[batch_no] = batch_list
	finally:
		driver.switch_to.window(main_window)
		batch_processed_count += 1
		firms_processed_count += len(batch_list)
		batch_time = time()	
		avg_time_per_batch = (batch_time-start_time)/\
				     batch_processed_count
		print "Averge time per batch: %.2f s" % avg_time_per_batch 
		print "====================================="

# Logout and close the browser
capiqLogout(driver, main_window)

""" Final print-out of summary statistics """
# Successes and failures
batch_successful_count = batch_processed_count - batch_failed_count
success_rate = 100.0*float(batch_successful_count)/float(batch_processed_count)

print "Processing of %d firms in %d batches completed"\
      % (firms_processed_count, len(download_list))
print "%d successful batches, %d failed batches, success rate: %.2f"\
      % (batch_successful_count, batch_failed_count, success_rate)

# Failed batches processing
failed_batches = sorted(failed_batches.keys())
if len(failed_batches) > 0:
	# Print out the failed batch numbers
	failed_batch_no = []
	for b_no in failed_batches:
		failed_batch_no.append(b_no)
	print "Failed batches: " + str(failed_batch_no)


	# Writing failed batch numbers to file 
	print "Writing batch numbers to failed_ids.txt"
	with open("failed_ids.txt", 'a') as fail_log:
		# Header with session and time
		session_end_time =  strftime("%a, %d %b %Y %H:%M +0800",\
				    localtime())
		fail_log.write("@ Session @ " + session_end_time + "\n")
		fail_log.write("File: " + argv[1] + " Report: " + argv[2] +\
			       "\n")

		# Write the failed batch numbers down 
		fail_log.write(str(failed_batch_no) + "\n\n")
			
# Print total Download time
total_download_time = time() - start_time
print "Total Download Time: %.0f min and %.2f sec" %\
      ((total_download_time/60), int(total_download_time)%60)


print "Script End"
