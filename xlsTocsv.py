import csv
import sys
import xlrd
import os
import getopt

#converts directory of xls files to csv files
#csv naming = 

def usage():
	print 'Usage : ' 
	print 'xlsTocsv.py -i directory'
	print 'with no argument current directory is default'
	sys.exit(2)

	
def csv_from_excel(dir, currXLS):
	wb = xlrd.open_workbook(os.getcwd() + dir + currXLS)
	
	#go through sheets
	for i in range(0, wb.nsheets):
		sh = wb.sheet_by_index(i)
		currCSV = open(currXLS +'_sh' +  str(i) + '.csv', 'wb')
		#each sheet, new csv file
		#TODO: merge sheets to one file
		wr = csv.writer(currCSV, delimiter =',', quotechar = '"', quoting = csv.QUOTE_ALL)
		
		for rownum in xrange(sh.nrows):
			currRow = []
			for col in range(sh.ncols):
				currCell = sh.cell_value(rownum,col)
				if type(currCell) == unicode:
					currCell = currCell.encode('utf-8', 'replace')
				currCell = str(currCell)
				currRow.append(currCell)
			wr.writerow(currRow)

		currCSV.close()
	
	
#TODO: make callable from other programs
def main(args):

	# arguments control
	dirname = ''
	
	try:
		opts, args = getopt.getopt(args,"hi:",["ifile="])
	except getopt.GetoptError:
		usage()
	for opt, arg in opts:
		if opt == '-h':
			usage()
		elif opt in ("-i", "--ifile"):
			dirname = arg
			
	cwd = os.getcwd()
	if len(dirname) > 0 and dirname[0] == '/':
		dirname = dirname[1:]
		
	
	inputFiles = os.listdir(dirname)
	
	excelFiles = []
	
	for file in inputFiles: 
		if file[-4:] == '.xls':
			excelFiles.append(file)
			
	
	#convert all .xls to .csv
	for file in excelFiles:
		csv_from_excel(dirname, file)

if __name__ == '__main__':
    main(sys.argv[1:])
