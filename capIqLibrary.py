from openpyxl import load_workbook
from collections import Counter
import xlrd

def getCompanyNamesInfo(workbook_path):


	# Load workbooks into a dictionary of {'firm_name' : [company id, batch_no]}
	master_table = load_workbook(workbook_path)
	ws = master_table.active
	print "Workbook %s loaded" % (workbook_path)

	for col_no in range(1, 20):
		col_title = ws.cell(row=1, column=col_no).value

		if col_title == "companyname":
			company_name_col = col_no
		elif col_title == "excelcompanyid":
			company_id_col = col_no
		elif col_title == "batch_no":
			batch_no_col = col_no

	print "columns: [names, ids, batch_no] = [%d, %d,  %d]"\
	      % (company_name_col,company_id_col, batch_no_col)

	company_names_info = {}
	for row_no in range(2,1000000):
		if ws.cell(row=row_no, column=1).value is None:
			break

		company_name = ws.cell(row = row_no, column = company_name_col).value
		batch_no = int(ws.cell(row = row_no, column = batch_no_col).value)
		company_id = ws.cell(row = row_no, column = company_id_col).value

		company_names_info[company_name] = [company_id, batch_no]

	print "%d company names and info loaded into memory" % (len(company_names_info))

	return company_names_info

# Get list of excel files and iterate through them
def getTrueName(rawfile, company_names_info):
	true_name = "Invalid"
	try:
		excel_file = xlrd.open_workbook(rawfile)
	except IOError:
		exit(rawfile + " not found.")
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

			try:
				excel_sheet = excel_file.sheet_by_name(sheet_name)
				company_name = excel_sheet.cell(1, 0).value

			# Where the data is out of alignment
			except IndexError:
				batch_no_vote.append(0)
				continue

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
				batch_no_vote.append(company_names_info[company_name][1])

			except IndexError:
				print "%s has no matching batch no." % (company_name)
				batch_no_vote.append(0)
			except KeyError:
				print "%s was not found in master list" % (company_name)
				batch_no_vote.append(0)

			# Once 4 firms are positively matched to a batch no, exit the loop 
			if len(batch_no_vote) >= 4:
				break
		
		# Find the most common batch_number
		batch_no = (Counter(batch_no_vote).most_common(1))[0][0]

		# If either report_type or zero data file, return Invalid
		if report_type == "unknown" or batch_no == 0:
			true_name = "Invalid"
		# Else, create true_name accordingly
		else:
			true_name = report_type + "_batch_" + str(batch_no) + ".xls"

	### Situations when the names returns invalid
	# 1. This file contains zero bytes
	except xlrd.biffh.XLRDError:
		true_name = "Invalid"
	# 2. The file could not be opened 
	except IOError:
		true_name = "Invalid"

	return true_name


def findMissing(downloaded_files, relations, last_batch):
	missing = []
	
	# Check if all files in the batch sequence exist
	# Record the missing batch numbers
	for no in range(1, last_batch+1):
		file_name = str(relations) + "_batch_" + str(no) + ".xls" 
		try:
			downloaded_files.index(file_name)
		except ValueError:
			missing.append(no)

	return missing


