from os import chdir, listdir
from sys import argv
from openpyxl import load_workbook
from collections import Counter
from capIqLibrary import getCompanyNamesInfo, getTrueName
import xlrd



# Initialize arguments
workbook_path = argv[1]
rawfiles_path = argv[2]
print ""
print "### Verify Capital IQ Buyer-Supplier Batch Files ####"
print "Raw files: %s" % (rawfiles_path)
print "Master file: %s" % (workbook_path)

company_names_info = getCompanyNamesInfo(workbook_path)
chdir(rawfiles_path)
raw_files = listdir(rawfiles_path)

for raw_file in raw_files:
	try:	
		true_name = getTrueName(raw_file, company_names_info)
		if not raw_file == true_name and not true_name == "unknown_batch_0.xls":
			print "Actual: %s" % (raw_file)
			print "Expect: %s" % (expected_filename)
			print ""
	except KeyError:
		continue

print "End of script"



