# capiq_crawler
The crawler extracts data from Capital IQ using the Screening and Report Builder function

Files
=====
0) all_gic_codes.txt
This is a list of all 157 GIC codes in Capital IQ. Copy relevant codes into a file named gic_codes_to_download.txt and run mass_screening.py to download data from Capital IQ.

1) capIqLibrary.py
This is a library containing commonly used routines for file-handling, file-reading and directory-checking

2) capIqNavigate.py
This library contains routines relating to the navigation of the Capital IQ website, including logging in, logging out,
inputting CIQ IDs into the Report Builder function, exporting data, inputting selection criteria into the Screening function,
setting the variable templates, choosing the firms which are being viewed and exporting the data, etc.

3) capiq_report.py
This file downloads Buyer-Supplier relations and corporate tree relations using the Report Builder function. It reads an .xlsx
files with the following columns (names, ids, batch_no), then pulls out appropriate information. It splits batches into smaller
pieces if they do not complete on the first try

4) example_dummy_file.xls
This file is used by capiq_report.py to stand-in for batches which produce 0 results

5) find_missing.py
This file is used to scan the downloads of capiq_report.py and mass_screening.py to identify gaps and missing batches and 
missing files in the download

6) mass_screening.py
This file exports data from Capital IQ in .xls files which contain a maximum of 10k firms each. It requires the user to first
create a template of selected variables in Capital IQ first using the name <num>_mass. It reads a .txt file to determine which
GIC codes it should filter by and subsequently download the data for.

7) verify_names.py
This file checks that the output of capiq_report.py. It ensures that the firms within match up to the batch no. given in the
.xlsx file containing the firms in that GIC code. It outputs files which have mis-matched firms.
