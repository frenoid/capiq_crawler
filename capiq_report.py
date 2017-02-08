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
from capIqLibrary import createDummyFile, getDownloadName, getBatchList, isDownloadDirClear, moveAllExcelFiles
from capIqAppendSubqueries import appendSubqueries
from random import shuffle

# Returns a list of batch numbers to download
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

	download_list.sort()
	print "Download type: %s" % (argv[2])
	print "Preparing to download %d batches" % (len(download_list))
	print "Batch # to download", str(download_list)
	print "***"

	return download_list

def renameBatchFile(batch_no, download_path, download_name, company_names_info):
	final_name = "Invalid"
	rename_success = False
	wait_time = 15

	# Check if the download is compelete by checking for .part files
	download_success = False
	while download_success is False:
		downloaded_files = listdir(download_path)
		download_success = True
		for downloaded_file in downloaded_files:
			if downloaded_file[-5:] == ".part":
				print "Download incomplete. Wait", str(wait_time), "seconds"
				download_success = False
				sleep(wait_time)
				break

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


def genSubqueries(batch_list, no_of_splits):	
	# Format {subquery_no : 'firm_ids'}
	sub_queries = {}
	sub_query_no = 0

	# Randomly re-sort the batch-list: to avoid name clashes
	shuffle(batch_list)
	no_of_firms_per_subquery = int(len(batch_list) / float(no_of_splits))

	# Create the sub_queries
	sub_start = 0
	while True:
		sub_query_no += 1
		sub_end = sub_query_no * no_of_firms_per_subquery

		# When reaching the end of the batch_list
		if sub_end >= len(batch_list):
			sub_end = len(batch_list) + 1
 
		# Create the sub_query
		sub_queries[sub_query_no] = batch_list[sub_start:sub_end]

		# If last sub_query has been created, break the loop
		if sub_end > len(batch_list):
			break
	 
		sub_start = sub_end

	return sub_queries

def subQuery(driver, batch_no, company_names_info, no_of_splits, download_id, download_path, code_name):
	success = True
	failed_ids = []
	print "Sub-querying batch %d" % (batch_no)
	
	batch_list = getBatchList(company_names_info, batch_no)
	sub_queries = genSubqueries(batch_list, no_of_splits)	
	print "Batch firm list split into %d sub-queries" % (len(sub_queries))
	
	for sub_query_no in sub_queries:
		try:
			print "++++++ Subquery #%d of batch #%d +++++++" % (sub_query_no, batch_no)
			driver.get("https://www.capitaliq.com/ciqdotnet/ReportsBuilder/CompanyReports.aspx")	
			print "Refresh Report Builder"
		
			sub_query_list = sub_queries[sub_query_no]
			print "Subquery #%d firm-list created" % (sub_query_no)

			# Add CQ IDs to the Report Generator
			valid_firm_count = addFirms(driver, sub_query_list)

			# Where there are no valid CQ IDs, create a dummy "No data" .xls file
			# Then proceed to next sub-query
			if valid_firm_count == 0:
				dummy_file_name = createDummyFile(batch_no, report_type)
				dummy_file_name = dummy_file_name[:(len(dummy_file_name)-4)] + "_" + str(sub_query_no) + ".xls"
				print "Dummy %s was created" % (dummy_file_name)
				print "Next batch"
				continue
	

			# Generate the report, min_wait_time is one-third of the full minimum wait time
			min_wait_time = (len(sub_query_list)/3.0)
			generateSuccess, download_name = generateReport(driver, batch_no, min_wait_time, download_id)

			if generateSuccess is True:
				# Rename the downloaded file, 3 tries allowed
				if download_name == "": 
					download_name = getDownloadName(report_type, valid_firm_count)

				rename_tries = 0
				rename_success = False
				while rename_tries < 5 and rename_success is not True:
					sleep(15)
					rename_tries += 1
					chdir(download_path)
					final_name, rename_success = renameBatchFile(batch_no, download_path,\
							             		     download_name, company_names_info)
					# Append the subquery no onto the batch_no
					if final_name is not "Invalid":
						sub_query_final_name = final_name[:(len(final_name)-4)]\
								       + "_" + str(sub_query_no) + ".xls"
						try:
							rename(final_name, sub_query_final_name)
							print "%s renamed to sub-query name %s"\
							       % (final_name, sub_query_final_name)
						except WindowsError:
							print "%s already exists. Rename failed"\
							      % (sub_query_final_name)


			
		except (TimeoutException, UnexpectedAlertPresentException,\
			NoSuchElementException, WebDriverException) as e:
			print "Exception", e, "encountered, sub-query incomplete"
			failed_ids.extend(sub_query_list)
			success = False
			continue

		finally:
			# Ensure focus is on main window
			# Moving excel files to classification folder
			final_path = download_path + code_name 
			files_moved = moveAllExcelFiles(download_path, final_path)
			print "%d excel files were moved to %s"\
			      % (files_moved, final_path)
			driver.switch_to.window(main_window)

	return success, failed_ids 


# 1: Check if there are enough arguments and if download directory is free of .xlsx files
download_path = "C:/Users/faslxkn/downloads/"
if (len(argv) < 5):
	print "%d arguments. Minimum is 4" % (len(argv)-1)
	print "Arguments: <file_of_ids> <{customer/supplier}] " +\
	      "<query_size> <start_batch> [end_batch]"
	exit()

if isDownloadDirClear(download_path) is False:
	exit("Download dir is not clear. Remove all .xls and .xlsx files")

# 2: Set report type
report_type = argv[2]
download_id = getReportType(report_type)

# 3: Read the file of firm IDs
code_name = argv[1]
company_names_info = getCompanyNamesInfo(code_name)
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
	driver, login_success = capiqLogin(driver, "davinchor@nus.edu.sg", "GPNm0nster")

	if login_success is False:
		print "Close browser. Wait one minute"
		sleep(60)
		driver.close()

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
keyboard_interrupt_flag = False

# Central loop: Terminates when the last batch is processed 
# while batch_processed_count < len(download_list):
for batch_no in download_list:
	try:
		batch_no = int(batch_no)
		print "++++++++++++Batch # %d++++++++++++++++" % (batch_no)
		batch_list = getBatchList(company_names_info, batch_no)

		print "Downloading batch #%d. To download %d of %d"\
		      % (batch_no, len(download_list), all_batch_count)

		""" # Break every 15 batches
		if batch_processed_count % 15 == 0 and batch_processed_count > 0:
			print "10 sec break"
			sleep(10)
		"""

		# Check if there have been more than 2 or more failures
		# Wait the equivalent number of minutes
		if consec_failure_count >= 2:
			print str(consec_failure_count), "consecutive failed downloads." 

		# If not first batch, then Refresh the RB to clear the Report Builder query list
		if batch_processed_count > 0:
			driver.get(report_page)	
			print "Refresh Report Builder"

		# Add CQ IDs to the Report Builder
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
		generateSuccess, download_name = generateReport(driver, batch_no, min_wait_time, download_id)

		if generateSuccess is True:
			if consec_failure_count > 0:
				print "Consecutive failures reset to 0"
				consec_failure_count = 0
			
			# If the download name is not returned by generateReport, manually create it instead
			if download_name == "": 
				download_name = getDownloadName(report_type, valid_firm_count)
				print "Guessed download_name:", download_name 

			# 10 tries allowed to rename the downloaded file, each time checking for partial downloads
			rename_tries = 0
			rename_success = False
			while rename_tries < 10 and rename_success is not True:
				sleep(10)
				rename_tries += 1
				chdir(download_path)
				final_name, rename_success = renameBatchFile(batch_no, download_path , download_name, company_names_info)
				if rename_tries == 10 and rename_success is not True:
					print "Rename tries:", str(rename_tries), "Max tries exceeded"

		# Where batch generation failed, throw TimeoutException
		else:
			raise TimeoutException("Batch generation failed")

	# Exception handling
	except (TimeoutException, UnexpectedAlertPresentException, NoSuchElementException, WebDriverException) as e:
		print "!Exception:", e, "proceed to next batch"
		consec_failure_count += 1
		print "Consecutive failed attempts: %d" % consec_failure_count
		batch_failed_count += 1
		failed_batches[batch_no] = batch_list
		continue

	except KeyboardInterrupt:
		print "!Exception: Keyboard interrupt"
		print "Script is cleaning-up. Please wait."
		keyboard_interrupt_flag = True
		break

	finally:
		# Moving excel files to classification folder
		final_path = download_path + code_name 
		files_moved = moveAllExcelFiles(download_path, final_path)
		print "%d excel files were moved to %s" % (files_moved, final_path)

		# Clean-up and report query and download time
		driver.switch_to.window(main_window)
		batch_processed_count += 1
		firms_processed_count += len(batch_list)
		batch_time = time()	
		avg_time_per_batch = (batch_time-start_time)/batch_processed_count
		print "Averge time per batch: %.2f s" % avg_time_per_batch 
		print "====================================="


""" Final print-out of summary statistics """
# Count the successes and success rate
batch_successful_count = batch_processed_count - batch_failed_count
success_rate = 100.0*float(batch_successful_count)/float(batch_processed_count)

print "Processing of %d firms in %d batches completed"\
      % (firms_processed_count, len(download_list))
print "%d successful batches, %d failed batches, success rate: %.2f"\
      % (batch_successful_count, batch_failed_count, success_rate)

failed_batches = sorted(failed_batches.keys())
print "Failed batches: ",
for failed_batch_no in failed_batches:
	print str(failed_batch_no),

print ""

# Subquering the failed batch numbers
if len(failed_batches) > 0 and keyboard_interrupt_flag is False:
	# Subquery each failed batch
	print "************************************"
	print "*           Subqueries             *"
	print "************************************"

remaining_ids = []
no_of_splits = 5

for failed_batch in failed_batches:
	subquery_success, subquery_failed_ids = subQuery(driver, failed_batch, company_names_info, no_of_splits,\
			                                 download_id, download_path, code_name)

	if subquery_success is True:
		print "Batch %d was resolved by subquery" % (failed_batch)
	else:
		print "Subquery batch %d failed" % (failed_batch)
		remaining_ids.extend(subquery_failed_ids)

# Logout and close the browser
capiqLogout(driver, main_window)

"""
# Print remaining failed batches into a text file
failed_batches = sorted(failed_batches.keys())
if len(failed_batches) > 0:
	# Print out the failed batch numbers
	failed_batch_no = []
	for b_no in failed_batches:
		failed_batch_no.append(b_no)
	print "Failed batches: " + str(failed_batch_no)


	# Writing failed batch numbers to file 
	print "Writing batch numbers to failed_batches.txt"
	with open("failed_batches.txt", 'a') as fail_log:
		# Header with session and time and report type
		session_end_time =  strftime("%a, %d %b %Y %H:%M +0800",\
				    localtime())
		fail_log.write("@ Session @ " + session_end_time + "\n")
		fail_log.write("File: " + argv[1] + " Report: " + argv[2] + "\n")

		# Write the failed batch numbers down 
		fail_log.write(str(failed_batch_no) + "\n\n")
"""

# Writing failed ids to file
print "Writing failed ids to failed_ids.txt"
with open("failed_ids.txt", 'a') as fail_log:
	# Header with session and time and report type
	session_end_time =  strftime("# %a, %d %b %Y %H:%M +0800", localtime())
	fail_log.write("# @ Session @ " + session_end_time + "\n")
	fail_log.write("# List of failed firm ids\n")
	fail_log.write("# File: " + argv[1] + " Report: " + argv[2] + "\n")

	for firm_id in remaining_ids:
		fail_log.write(firm_id + "\n")
	
# Print total Download time
total_download_time = time() - start_time
print "Total Download Time: %.0f min and %.2f sec" %\
      ((total_download_time/60), int(total_download_time)%60)

print "Script End"
