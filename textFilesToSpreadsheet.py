#! python3
# textFilesToSpreadsheet.py - Reads the content of text file(s),
# and inserts content into new spreadsheet.

import openpyxl, sys

if len(sys.argv) > 1:

	wb = openpyxl.Workbook()
	sheet = wb.active

	files = sys.argv[1:]

	# Each file will be stored as its own column. (x, y) = (c, r)
	for c in range(len(files)):
	
		try:
			location = str(files[c])
		
		except Exception as e:
			print(e)

		text_file = open(files[c])

		lines = text_file.readlines()

		text_file.close()

		# Each line will be saved as its own row.
		for r in range(len(lines)):
			sheet.cell(row = r + 1, column = c + 1).value = lines[r]

		print(files[c] + ' was successfully read.')

	# Use the first filename given in saving new filename.
	p = files[0][:-4] + '_text_files.xlsx'
	
	wb.save(p)
	
	print('Text files successfully saved to spreadsheet as ' + p)

else:
	print('You must include file name(s) in your argument.')