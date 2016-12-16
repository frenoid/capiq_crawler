from os import chdir, listdir
from sys import argv
from openpyxl import load_workbook
from collections import Counter
import xlrd

def getCompanyNamesBatchNo(workbook_path):


	# Load workbooks into a dictionary of {'firm_name' : batch_no}
	master_table = load_workbook(workbook_path)
	ws = master_table.active
	print "Workbook %s loaded" % (workbook_path)

	for col_no in range(1, 20):
		col_title = ws.cell(row=1, column=col_no).value

		if col_title == "companyname":
			company_name_col = col_no
		elif col_title == "batch_no":
			batch_no_col = col_no

	print "Company names in col %d, IDs in col %d" % (company_name_col, batch_no_col)

	company_names_batch_no = {}
	for row_no in range(2,1000000):
		if ws.cell(row=row_no, column=1).value is None:
			break

		company_name = ws.cell(row = row_no, column = company_name_col).value
		batch_no = int(ws.cell(row = row_no, column = batch_no_col).value)

		company_names_batch_no[company_name] = batch_no

	print "%d company names loaded into memory" % (len(company_names_batch_no))

	return company_names_batch_no

# Get list of excel files and iterate through them
def getExpectedName(rawfile, company_names_batch_no):
	try:
		excel_file = xlrd.open_workbook(rawfile)
		sheet_names = excel_file.sheet_names()
		report_type = "unknown"
		batch_no_vote = [] 

		# Iterate across the sheets to guess the batch no
		# Stop when either 4 firms are matched or there are no worksheets left
		for sheet_name in sheet_names:
			# No data => next sheet
			if sheet_name == "No Data":
				batch_no_vote.append(0)
				continue

			excel_sheet = excel_file.sheet_by_name(sheet_name)
			company_name = excel_sheet.cell(1, 0).value

			# No data => next sheet
			if company_name == "No Data":
				batch_no_vote.append(0)
				continue

			# Get the company name and report type
			company_name = company_name[:(company_name.index('>') - 1)]
			report_type = excel_sheet.cell(3, 0).value
			if report_type == "  Customers":
				report_type = "customers"
			elif report_type == "  Suppliers":
				report_type = "suppliers"

			# Get the batch_no
			try:
				batch_no_vote.append(company_names_batch_no[company_name])

			except IndexError:
				print "No matching batch no, for %s" % (company_name)
				batch_no_vote.append(0)

			# Once 4 firms are positively matched to a batch no, exit the loop 
			if len(batch_no_vote) >= 4:
				break
		
		# Find the most common batch_number
		batch_no = (Counter(batch_no_vote).most_common(1))[0][0]
		expected_filename = report_type + "_batch_" + str(batch_no) + ".xls"


	# This file contains zero bytes
	except xlrd.biffh.XLRDError:
		expected_filename = "zero_byte_file"

	return expected_filename

""" Main() """
# Initialize arguments
workbook_path = argv[1]
rawfiles_path = argv[2]
print ""
print "### Verify Capital IQ Buyer-Supplier Batch Files ####"
print "Raw files: %s" % (rawfiles_path)
print "Master file: %s" % (workbook_path)

company_names_batch_no = getCompanyNamesBatchNo(workbook_path)
chdir(rawfiles_path)
raw_files = listdir(rawfiles_path)

for raw_file in raw_files:
	try:	
		expected_filename = getExpectedName(raw_file, company_names_batch_no)
		if not raw_file == expected_filename and not expected_filename == "unknown_batch_0.xls":
			print "Actual: %s" % (raw_file)
			print "Expect: %s" % (expected_filename)
			print ""
	except KeyError:
		continue

print "End of script"

