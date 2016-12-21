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
from capIqLibrary import getCompanyNamesInfo, getTrueName
from time import sleep, time, localtime, strftime
from os import chdir, remove, rename, listdir
from shutil import copy
from math import ceil
from capIqNavigate import getReportType, capiqInitialize, capiqLogin, getValidFirmCount, addFirms, generateReport, capiqLogout
from capIqLibrary import createDummyFile, getDownloadName


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


def getDownloadList(company_names_info, argv):
	query_size = int(argv[3])
	query_type = argv[4]
	batch_total = int(ceil(len(company_names_info)/float(query_size)))

	# Calculate the number of batches to download 
	print "%d firms with query-size of %d makes %d batches"\
	      % (len(company_names_info), query_size, batch_total)
	
	print "Step 2: Selecting firms to download"

	# Arg4 can be either "all", "list", or an integer
	download_list = []

	# Get all batches 
	if(query_type == "all"):
		download_list = range(1, batch_total+1)
	
	# Get a list of batches
	elif(query_type == "list"):
		download_list.extend(argv[5:]) 
		for download_batch in download_list:
			download_batch = int(download_batch)
			if(download_batch > batch_total):
				print "Batch #%d exceeds total %d"\
				      % (download_batch, batch_total)
				exit("Batch exceeds batch range")

	# Get a range of batches if arg4 is an integer
	elif(query_type > 0):
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

	print "Download type: %s" % (argv[2])
	print "Preparing to download %d batches" % (len(download_list))
	print download_list
	print "***"

	return download_list

def getBatchList(company_names_info, batch_no):
	
	# Batch creation
	batch_list = []

	for company in company_names_info:
		if company_names_info[company][1] == batch_no:
			firm_id = company_names_info[company][0]
			batch_list.append(firm_id)

	print "Batch #%d has been created" % (batch_no)

	return batch_list


	return int(count)

def renameBatchFile(download_files, download_name, company_names_info):
	rename_success = False

	for index in range(len(downloaded_files)):
		actual_name = str(downloaded_files[index])

		## Use 2 methods to generate the filename, if they agree, rename the file
		## if not use batch no to rename the file, but mark the file with "star_"

		# method 1: matching firm name with master table
		if actual_name == download_name:
			true_name = getTrueName(actual_name, company_names_info)

		# method 2: check download batch no and file name
			if actual_name[-13:] == "Suppliers.xls":
				batch_filename = "suppliers_batch_" + str(batch_no) + ".xls"
			elif actual_name[-13:] == "Customers.xls":		
				batch_filename = "customers_batch_" + str(batch_no) + ".xls"

			try:
				# Where the 2 methods agree
				if true_name == batch_filename or true_name is "Invalid":
					rename(actual_name, batch_filename)
					print "Batch agrees with master, File renamed to %s"\
					      % (batch_filename)

				# Where the 2 methods disagree
				else:
					rename(actual_name, "star_" + batch_filename)
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

	return rename_success




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
download_list = getDownloadList(company_names_info, argv)
all_batch_count = int(ceil(len(company_names_info)/float(batch_size)))

# 4: Initialize the brower and load Capital IQ 
# Allow 3 attempts before closing the browser
login_attempts = 0
login_success = False
while login_attempts < 3 and login_success is False:
	login_attempts += 1	
	report_page = "https://www.capitaliq.com/ciqdotnet/ReportsBuilder/CompanyReports.aspx"
	driver = capiqInitialize(report_page)
	main_window = driver.current_window_handle
	print "Capital IQ website loaded"
	print "Login attempt #%d" % (login_attempts)
	login_success = capiqLogin(driver, "davinchor@nus.edu.sg", "GPNm0nster")
	if login_success is False:
		sleep(60)

if login_attempts == 3:
	exit("Login attempts limit exceeded.")

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
		batch_list = getBatchList(company_names_info, batch_no)

		print "Downloading batch #%d. To download %d of %d"\
		      % (batch_no, len(download_list), all_batch_count)

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
		download_name, generateSuccess = generateReport(driver, batch_no, min_wait_time, download_id)

		if generateSuccess is True:
			if consec_failure_count > 0:
				print "Consecutive failures reset to 0"
				consec_failure_count = 0

			
			# Rename the downloaded file, 3 tries allowed
			if download_name == "": 
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
