from sys import argv
from os import chdir, listdir
from capIqLibrary import findMissing

# Initiate arguments
download_folder = argv[1]
relations = argv[2]
last_batch = int(argv[3])

chdir(download_folder)
downloaded_files = listdir(download_folder)
downloaded_files.sort()
customers_missing = []
suppliers_missing = []

print ""
print "*** Find missing batch no in Cap IQ rawfiles ***"
print "Check for %s missing batches in %s" % (relations, download_folder)
print "Last batch no: %d" % (last_batch)

# Check if all files in the batch sequence exist
# Record the missing batch numbers
if relations == "all" or relations == "customers":
	customers_missing = findMissing(downloaded_files, "customers", last_batch)

if relations == "all" or relations == "suppliers":
	suppliers_missing = findMissing(downloaded_files, "suppliers", last_batch)

# Print out the missing batch numbers
print "Missing customers: " + str(customers_missing)
print ""
print "Missing suppliers: " + str(suppliers_missing)

print "End of program"

