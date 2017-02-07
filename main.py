import xml.etree.ElementTree as et
import xlsxwriter as xl
import  argparse

def parseDate(date):
	list = date.split('-')
	year = int(list[0])
	month = int(list[1])
	day_time = list[2].split('T')
	day = int(day_time[0])
	time = day_time[1].split('.')[0]
	return year, month, day, time

parser = argparse.ArgumentParser(description="Tool to convert evtxexport, evtxtract XML output into XLSX format")
parser.add_argument("input",  help="Path to events XML file")
parser.add_argument("output", help="Path to XLSX output file")
parser.add_argument("--debug", "-d", dest="debug", action="store_true", help="Show Debugging Information")
args = parser.parse_args()

xmlEvtPath = args.input
output = args.output

if args.debug:
	print "[+] Parsing XML Event Log File"
all = et.parse(xmlEvtPath)
events = all.getroot()

wb = xl.Workbook(output)
ws = wb.add_worksheet("AllEvents")
ws.write(0,0, "Event Record ID")
ws.write(0,1, "Year")
ws.write(0,2, "Month")
ws.write(0,3, "Day")
ws.write(0,4, "Time")
ws.write(0,5, "Event ID")
ws.write(0,6, "Computer")
#try:
for i in range(1, len(events)):
	if args.debug:
		print "[+] Parsing Event Record", events[i][0][8].text
	#Filling EventRecordID, TimeCreated, EventID, Computer
	ws.write(i,0,int(events[i][0][8].text)) # EventRecordID
	year, month, day, time = parseDate(events[i][0][7].attrib['SystemTime']) # SystemTime
	ws.write(i,1,year)
	ws.write(i,2,month)
	ws.write(i,3,day)
	ws.write(i,4,time)
	ws.write(i,5,int(events[i][0][1].text)) # EventID
	ws.write(i,6,events[i][0][12].text) # Computer

	#Fililng Event
	if len(events[i][1]) != 0:
		for j in range(0, len(events[i][1])):
			if events[i][1].tag == "EventData":
				try:
					record = str(events[i][1][j].attrib['Name']) + '=' + str(events[i][1][j].text)
				except UnicodeEncodeError:
					record = "Encoding Error"
				ws.write(i,j + 7, record)
'''except:
	if args.debug:
		print "[-] Error"
	wb.close()
'''
wb.close()

