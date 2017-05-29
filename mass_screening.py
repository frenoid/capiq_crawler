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
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep, time, localtime, strftime
from os import chdir, remove, rename, listdir
from shutil import copy
from math import ceil, floor
from capIqNavigate import capiqInitialize, capiqLogin, capiqLogout, generateReport, downloadFile
from capIqLibrary import createDummyFile, getDownloadName,isDownloadDirClear, moveAllExcelFiles, moveAllPartialFiles, readDownloadDir, checkMakeDir, checkDownloadComplete

# Uses the old company screening interface
def switchToOldScreening(driver):
	try:
		old_switch = WebDriverWait(driver, 30).until(
			     EC.presence_of_element_located((By.ID,"returnToOriginalLinkNonBeta"))
			     )
		old_switch.click()
		sleep(3)
	except (TimeoutException, NoSuchElementException):
		driver.quit
		exit("!Exception at switching to old interface")

	return 

# Get the ScreenId, a number given to each Company Screen performed
def getScreenId(browser_url):
	# Get the relevant string in the url
	begin_i = browser_url.find("UniqueScreenId=")
	end_i = browser_url.find("&", begin_i)
	screen_url = browser_url[begin_i:end_i]

	# Extract the numbers in the url to get the screen id
	screen_id = filter(lambda x: x.isdigit(), screen_url)
	print "Screen ID is %s" % (screen_id)

	return screen_id

# This examines the Company Screening page to figure out which block of 50,000 firms you are currently on
def getPageNo(driver):
        assert "Company Screening" in driver.title
        page_no = 0

        page_no_textline = WebDriverWait(driver, 30).until(
                           EC.presence_of_element_located((By.XPATH, "/html/body/table/tbody/tr[2]/td[4]/div/form/div[6]/div[4]/div[2]/table/tbody/tr/td/span[1]/table[3]/tbody/tr[2]/td/nobr")) 
                           ).text

        begin_i = page_no_textline.find("of ")
        end_i = page_no_textline.find(" to", begin_i)
        if end_i == -1:
            page_no = 1
        else:
            first_firm = int(page_no_textline[begin_i+3:end_i])
            page_no = int(((first_firm-1) / float(50000)) + 1)

        return page_no

# Set the correct GIC code to filter firms
def setGicFilter(driver, gic_code):
	# Enter the GIC code in the search box
	screening_search = WebDriverWait(driver, 30).until(
			   EC.presence_of_element_located((By.ID, "SearchDataPointsAutoCompleteTextBoxPhase2"))
			   )

	screening_search.send_keys(gic_code)
	sleep(1)
	# screening_search.send_keys(Keys.ENTER)

	# Click on the first result, 
        # Communications Equipment and Electronic Components chooses the second result
	sleep(5)
        if gic_code == "Communications Equipment" or gic_code == "Electronic Components":
            print "%s -> Select second search result" % (gic_code)
            sub_search = WebDriverWait(driver,15).until(
		         EC.presence_of_element_located((By.XPATH,\
                         "/html/body/table/tbody/tr[2]/td[4]/div/form/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td/div/span/div/div[2]/div[2]/a/div[1]/span/b/span"))
		         )

        else:
	    sub_search = WebDriverWait(driver,15).until(
		         EC.presence_of_element_located((By.XPATH,\
                         "/html/body/table/tbody/tr[2]/td[4]/div/form/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td/div/span/div/div[1]/div[2]/a"))
		         )
	sub_search.click()

        # Choose primary code only
	primary_only = WebDriverWait(driver,15).until(
		       EC.element_to_be_clickable((By.XPATH,\
		       "/html/body/table/tbody/tr[2]/td[4]/div/form/div[3]/table/tbody/tr/td/table[1]/tbody/tr/td/div/span/div/div[1]/a/span/b"))
		       )
	primary_only.click()
	sleep(3)
	screening_search.send_keys(Keys.ENTER)
	print "Criterion added"

	# Scroll to the bottom of the page
	sleep(3)
	driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")


	# Note that the following occurs in a different browser frame
	# In the results frame, get the number of firms, the view results
	driver.switch_to_frame("CriterionResultsFrame")
	firm_count_elem = WebDriverWait(driver,15).until(
			  EC.element_to_be_clickable((By.ID,
		          "_viewTopControl__numberOfResults"))
			  )

	# Get the number firms in the GIC code
	firm_count = int(filter(lambda x: x.isdigit(), firm_count_elem.text))
	screen_id = getScreenId(driver.current_url)

	view_results = WebDriverWait(driver,15).until(
			EC.element_to_be_clickable((By.ID,
			"_viewTopControl__resultsLink"))
			)
	view_results.click()

	return firm_count, screen_id

# Sets the correct template to get the correct variables
def setTemplate(driver, option):
	sleep(3)
	option = int(option)
	template_name = "Invalid"

	# We are not in the right page
	if "Screening Results" not in driver.title:
		print("!Exception. Not in screening page")
                return "Invalid"

	# Check if the option is a digit, if not exit with error
	if(str(option).isdigit() == False):
		print "Invalid template no:" , str(option)
                return "Invalid"

	# Get the drop down menu to enter template name
	drop_down = WebDriverWait(driver,15).until(
		    EC.element_to_be_clickable((By.ID,
		    "_displayOptions_Displaysection1_SelectedTemplate"))
		    )
	drop_down.click()
	sleep(3)
	template_name = str(option) + "_mass"
	drop_down.send_keys(template_name)
        sleep(3)

	# Click the set template button
	set_template_go=WebDriverWait(driver,15).until(
	     	     EC.element_to_be_clickable((By.ID,
	   	     "_displayOptions_Displaysection1_GoButton"))
		     )
        try:
	    # 2 clicks, one to set focus, one to actually click on the Go button	
	    set_template_go.click()
            set_template_go.click()
	    sleep(5)
        except WebDriverException:
            pass
        finally:
	    # Wait till template is fully loaded
	    WebDriverWait(driver,120).until(
	    EC.element_to_be_clickable((By.ID,
	    "_displayOptions_Displaysection1_ReportingOptions_GoButton"))
	    )

	return template_name

# Choose the correct set of 10,000 firms to export
def chooseExportFirms(driver, download_no):

	# Switch to correct set of 0 - 50000 firms
	# Using the Export: dropdown menu
	export_menu = WebDriverWait(driver, 30).until(
		      EC.element_to_be_clickable((By.ID,
		      "_displayOptions_Displaysection1_ReportingOptions_NumberOfTargetsExcel"))
		      )
	export_menu.click()

        if download_no % 5 == 1:
                export_menu.send_keys("Top 10000")
	elif download_no % 5 == 2:
	        export_menu.send_keys("10001")
	elif download_no % 5 == 3:
		export_menu.send_keys("20001")
	elif download_no % 5 == 4:
		export_menu.send_keys("30001")
	elif download_no % 5 == 0:
		export_menu.send_keys("40001")
	export_menu.send_keys(Keys.ENTER)
	
	print "Switched to %d th firms" % ((download_no-1)*10000)

	return

# Change to the correct set of 10,000 or 50,000 firms
def changePageNo(driver, download_no, screen_id):

        # Get the current page
	assert "Screening Results" in driver.title
        current_page = getPageNo(driver)

        # Get expected page no
        page_no = int(ceil(float(download_no)/float(5)))

        # Compare current page to expected page_no
        # If not the same, change page
        if current_page == page_no:
                print "Current page %d and correct" % (current_page)
        else:
                print "Changing page %d -> page %d" % (current_page, page_no)

		# Return to the screen selection page
		screen_url = "https://www.capitaliq.com/CIQDotNet/Screening/ScreenBuilder.aspx?UniqueScreenId=" + screen_id + "&clear=all&returntooriginal=1#"
		driver.get(screen_url)

		# Switch to the correct set of 50,000 firms
		sleep(10)
		driver.switch_to_frame("CriterionResultsFrame")
		view_range_menu = WebDriverWait(driver, 15).until(
				  EC.element_to_be_clickable((By.ID,
				  "_viewTopControl__range"))
				  )

		# Set view-range by selecting the first firm no
		first_firm_no = (download_no * 10000) - 9999
		view_range_menu.send_keys(str(first_firm_no))

		# View results
		view_results = WebDriverWait(driver,15).until(
			       EC.element_to_be_clickable((By.ID,
			       "_viewTopControl__resultsLink"))
			       )
		view_results.click()

	# Use the export menu to select the correct group of 10,000 firms
	chooseExportFirms(driver, download_no)


	return

# Rename the downloaded file according to its GIC code and page number
# File name format <gic_code>_<page_no)_of_<total_pages>.xls
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

# 0: Program starts
# Check if there are enough arguments passed in
if(argv[1] == "--help"):
	print "Help for Capital IQ mass screening downloader"
	print ""
	print "Format: python mass_screening.py <GIC_CODE> <TEMPLATE_NO> <DOWNLOAD_TYPE> [FILE_NO]"
	print "DOWNLOAD_TYPE: all list"
	exit()

print "********** Capital IQ mass screening downloader *********"
if (len(argv) < 3):
	print "Format: python mass_screening.py <GIC_CODE> <TEMPLATE_NO> <DOWNLOAD_TYPE> [FILE_NO]"
	print "DOWNLOAD_TYPE: all list"
	exit("!Error: Insufficient arguments")

# Parse arguments in program variables, check if they are valid
target_gic = argv[1]
template_no = argv[2]
download_type = argv[3]
if download_type != "all" and download_type != "list":
    print "%s is not a valid download type" % download_type
    exit("!Error: Invalid download type, check your arguments")

# 1: Ensuring that download_dir exists and is clear
# Ensure that final storage folder also exists
download_path = readDownloadDir("download_dir.txt")
final_path = download_path + "/" + target_gic
checkMakeDir(final_path)

# Check if download_path exists
if download_path == "Invalid":
        print "Download path not found"
        exit("!Error: Invalid download directory")

# If not, clear it by moving all excel files to the final_dir
# Also move all .part files to the final dir
while isDownloadDirClear(download_path) is False:
        print "Download dir %s is not clear" % (download_path)
        excel_files_moved = moveAllExcelFiles(download_path, final_path)
        print "%d .xls files moved" % (excel_files_moved)
	partial_files_moved = moveAllPartialFiles(download_path, final_path)
	print "%d .part files moved" % (partial_files_moved)

# Login, set GIC filter and template until successful, 5 tries allowed
initiate_success = False
initiate_attempts = 0
while initiate_success != True and initiate_attempts < 5:
        initiate_attempts += 1
        print "Initialization attempt #", str(initiate_attempts)

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

	if login_attempts > 3:
                print "Max logins reached"
                continue

	# Wait one minute for the screening page to fully load
	wait_time = 0
        screening_page_loaded = False
	while screening_page_loaded != True:
		if "Company Screening" in driver.title:
			print "Company Screening loaded"
                        screening_page_loaded = True
		elif wait_time > 60:
			print("Timeout on loading Company Screening page")
                        break
		else:
			sleep(5)
			wait_time += 5
        if screening_page_loaded == False:
                print "!Exception, screening page not loaded"

	# 3. Set GIC filter and variable template
	try:
		# Set filter to target GIC code and get number of firms
		total_firm_count, screen_id = setGicFilter(driver, target_gic)
		print "Filter set: %d firms in %s" % (total_firm_count, target_gic)
	except(TimeoutException, NoSuchElementException, UnexpectedAlertPresentException):
		driver.quit()
		print "!Exception while setting GIC filter"
                continue

	try:
		# Set correct variable template
		template_name = setTemplate(driver, template_no)
		if template_name == "Invalid":
			driver.quit()
			print "Invalid template name"
                        continue
		print "Template set: %s" % (template_name)
	except(TimeoutException, NoSuchElementException, UnexpectedAlertPresentException, WebDriverException):
		driver.quit()
		print "!Exception while setting template"
                continue

        initiate_success = True
        print "* Initialization successful *"

# 4. Create download list, one file for every 10,000 firms
total_files = int(ceil(float(total_firm_count)/10000.0))

if download_type == "all":
    download_list = range(1, total_files+1)
elif download_type == "list":
    download_list = map(int, argv[4:])
else:
    print "%s is not a valid download type" % (download_type)
    driver.quit()
    exit("Invalid download type. Check your arguments")


# 5. Main loop, commence download
# Iterate across files. Each file containing up to 10,000 firms
failed_page_downloads = []
sleep(15)
print "Download list: %s" % (str(download_list))
print "***** Download Commenced *****"

for download_no in download_list:
	print "=== File %d of %d ===" % (download_no, total_files)
        download_success = False
        download_attempts = 0

	# The download will re-try until it succeeds or 
	# the max attemps is exceed
        while download_success != True and download_attempts < 5:
            download_attempts += 1
            print "* Attempt #%d *" % (download_attempts)
	    try:
	        # Change to approriate page number
	    	changePageNo(driver, download_no, screen_id)	

		# Initiate download, allow max of 6 minutes for download to generate
		min_wait_time = 10
		max_wait_time = 360
		success, filename = generateReport(driver, 0, min_wait_time,\
					    max_wait_time, "_displayOptions_Displaysection1_ReportingOptions_GoButton")

		# Ensure the download is done	
		sleep(5)
		download_complete = False
		total_wait_time = 0
		print "Downloading. Time elapsed:",
		while download_complete == False and total_wait_time < 180:
			download_complete = checkDownloadComplete(download_path)
			if download_complete == False:
				sleep(10)
				print str(total_wait_time),
				total_wait_time += 10

		# Rename to download file accordingly
		renameMassFile(download_path, filename, target_gic,download_no, total_files)

		# If the above code executes without exceptions, then flag the download as successful
		# This flag is required to exit the download loop
		download_success = True

	    #If exception, return to screening page
	    except(TimeoutException, NoSuchElementException,\
	       UnexpectedAlertPresentException,StaleElementReferenceException) as exception_type:
	        print "!Exception of type", str(exception_type), "encountered"
		driver.switch_to_window(main_window)
		driver.get(driver.current_url)
		sleep(10)
						
	    finally:
		excel_files_moved = moveAllExcelFiles(download_path, final_path)
		print "%d .xls files moved" % (excel_files_moved)
		partial_files_moved = moveAllPartialFiles(download_path, final_path)
		print "%d .part files moved" % (partial_files_moved)
		driver.switch_to_window(main_window)

            # If the download did not complete successfully, mark as such
	    if download_success != True:
		failed_page_downloads.append(download_no)


print "Failed pages:", str(failed_page_downloads)
capiqLogout(driver, main_window)
print "Script End"
