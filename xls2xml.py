import xlrd, sys, traceback,re,os
import HTMLParser
from xml.dom import minidom
reload(sys)
sys.setdefaultencoding( "utf-8" )

def open_xls(file):
	book = xlrd.open_workbook(file)
	return book

def get_sheets_info(book):
	sheetNum = len(book.sheets())
	sheetNames = []
	for sheetIndex in range(sheetNum):
		sheet_name=book.sheet_names()[sheetIndex]
		sheetNames.append(sheet_name)
	return sheetNum,sheetNames

def get_sheet_data(book, sheetIndex = 0, colNameIndex = 0):
	sheet = book.sheets()[sheetIndex]
	nrows = sheet.nrows
	ncols = sheet.ncols
	colNames =  sheet.row_values(colNameIndex)	
	data = []
	for rowIndex in range(1,nrows):
		row = sheet.row_values(rowIndex)
		if row:
			item = {}
			for i in range(len(colNames)):
				item[colNames[i]] = row[i]
			data.append(item)
	return data

def write_xml(data, xmlFile):
	try:
		f = open(xmlFile,'w')
		
		try: 
			doc = minidom.Document() 
			rootNode = doc.createElement("testcases") 
			doc.appendChild(rootNode)

			for i in range(len(data)):
				caseNode = doc.createElement("testcase")
				caseNode.setAttribute("name", data[i]['name']) 
				rootNode.appendChild(caseNode) 

				preconditionsNode = doc.createElement("preconditions") 
				caseNode.appendChild(preconditionsNode)
				preconditionsText = '<![CDATA[<p>'+re.sub('\s','</p><p>',data[i]['preconditions'])+'</p>]]>'
				preconditionsTextNode = doc.createTextNode(preconditionsText)
				preconditionsNode.appendChild(preconditionsTextNode)		

				stepsNode = doc.createElement("steps")
				caseNode.appendChild(stepsNode)
				stepsText = '<![CDATA[<p>'+re.sub('\s','</p><p>',data[i]['steps'])+'</p>]]>'
				stepsNodeTextNode = doc.createTextNode(stepsText)
				stepsNode.appendChild(stepsNodeTextNode)

				expectedresultsNode = doc.createElement("expectedresults")
				caseNode.appendChild(expectedresultsNode)
				expectedresultsText = '<![CDATA[<p>'+re.sub('\s','</p><p>',data[i]['expectedresults'])+'</p>]]>'
				expectedresultsNodeTextNode = doc.createTextNode(expectedresultsText)
				expectedresultsNode.appendChild(expectedresultsNodeTextNode)

			doc.writexml(f, "", "\t", "\n", "utf-8")

		except: 
			trackback.print_exc() 
	except: 
		print "open file failed" 
	finally: 
		f.close() 



def unescape_xml(xmlFile):
	html_parser = HTMLParser.HTMLParser()
	try:
		xf = open(xmlFile,'r+')
		rstr = xf.read().decode("utf-8")
		ustr = html_parser.unescape(rstr)
		#print ustr
		xf.truncate(0)
		xf.seek(0)
		xf.write(ustr.encode('utf-8'))
	except: 
		print "sth wrong." 
	finally: 
		xf.close() 



def main():
	savePath = os.path.split(sys.argv[1])[0]

	xlsBook = open_xls(sys.argv[1])
	sheet_num,sheet_names = get_sheets_info(xlsBook)
	for i in range(sheet_num):
		data = get_sheet_data(xlsBook,i)
		xmlPath = savePath+"/"+sheet_names[i]+".xml"
		write_xml(data, xmlPath)
		unescape_xml(xmlPath)

if __name__=="__main__":
	main()