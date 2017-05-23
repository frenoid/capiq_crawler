from sys import argv
from os import chdir, listdir
from capIqLibrary import findMissing

# Finds missing batch numbers for the following download types:
# Customers, Suppliers, Corporate Tree
def getReportRelations(download_folder, relations):
    last_batch = int(argv[3])
    chdir(download_folder)
    downloaded_files = sorted(listdir(download_folder))
    customers_missing, suppliers_missing = [], []

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

    if relations == "corporate_tree":
        corptree_missing = findMissing(downloaded_files, "corporateT", last_batch)

    # Print out the missing batch numbers
    if len(customers_missing) > 0:
        print "Missing customers: " + str(customers_missing)
        return customers_missing

    if len(suppliers_missing) > 0:
        print "Missing suppliers: " + str(suppliers_missing)
        return suppliers_missing

    if len(corptree_missing) > 0:
        print "Missing corporate_tree: " + str(corptree_missing)
        return corptree_missing

# Find missing files for downloads from Capital IQ's Company Screening function
def getScreeningRelations(download_folder, relations):
    chdir(download_folder)
    gic_code_folders = listdir(download_folder)
    gic_code_folders = sorted(gic_code_folders)
    missing_files = []

    # Check if the download folders exist and that there are 157 of them
    if gic_code_folders is None:
        print "No GIC codes found"
        return
    elif len(gic_code_folders) < 157:
        print "Incomplete download. Only %d GIC codes present" % len(gic_code_folders)
    else:
         print "There are %d GIC codes present" % len(gic_code_folders)

    # Iterate across GIC codes, check if all files are present in each GIC code
    for gic_folder in gic_code_folders:
        # print "Checking %s" % (gic_folder)
        try:
            chdir(download_folder)
            raw_files = listdir(gic_folder)
        except WindowsError:
            print gic_folder, "is not a valid directory"
            continue
        chdir(gic_folder)

        try:
            total_no_files = int(filter(lambda x: x.isdigit(), raw_files[0][-6:-4]))
        except IndexError:
            print "* No files found in %s" % gic_folder
            continue

        # print "%s should have %d files" % (gic_folder, total_no_files)

        # Generate list of expected files, mark those which don't appear
        for file_no in range(1, total_no_files+1):
            expected_file = gic_folder + "_" + str(file_no) + "_of_" + str(total_no_files) + ".xls"
            is_file_missing = True
            for raw_file in raw_files:
                if raw_file == expected_file:
                    is_file_missing = False 
                    break

            if is_file_missing == True:
                missing_files.append(expected_file)

    print "%d files are missing" % (len(missing_files))
    print "===================="
    for missing_file in missing_files:
        print "*", missing_file

    return missing_files




if __name__ == "__main__":
    # Initiate arguments
    download_folder = argv[1]
    relations = argv[2]

    if relations == "all" or relations == "customers" or relations == "suppliers" or relations == "corporateT":
        getMissingReportRelations(download_folder, relations)
        
    elif relations  == "screening":
        getScreeningRelations(download_folder, relations)
        


