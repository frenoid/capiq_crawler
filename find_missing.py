from sys import argv
from os import chdir, listdir

# Initiate arguments
download_folder = argv[1]
relations = argv[2]
last_batch = int(argv[3])

downloaded_files = listdir(download_folder)
customer_missing = []
supplier_missing = []

print ""
print "*** Find missing batch no in Cap IQ rawfiles ***"
print "Check for %s missing batches in %s." % (relations, download_folder)
print "Last batch no: %d" % (last_batch)

# Check if all files in the batch sequence exist
# Record the missing batch numbers
if relations == "all" or relations == "customer":
	for no in range(1, last_batch+1):
		file_name = "customers_batch_" + str(no) + ".xls" 
		try:
			downloaded_files.index(file_name)
		except ValueError:
			customer_missing.append(no)

if relations == "all" or relations == "supplier":
	for no in range(1, last_batch+1):
		file_name = "suppliers_batch_" + str(no) + ".xls" 
		try:
			downloaded_files.index(file_name)
		except ValueError:
			supplier_missing.append(no)

# Print out the missing batch numbers
print "Missing customers: " + str(customer_missing)
print ""
print "Missing suppliers: " + str(supplier_missing)

print "End of program"

