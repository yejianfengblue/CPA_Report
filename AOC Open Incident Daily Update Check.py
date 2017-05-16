from win32com.client import Dispatch, constants as const
from datetime import datetime, date
import logging, os, time
from bs4 import BeautifulSoup as Soup

DebugMode = False
today = date.today()
workDir = "C:\Users\IBM_ADMIN\Desktop\CPA-AOC\Daily Report INC 1600\\"
mailCCList = ["Lei Wang", "Ce Zheng", "whtse@hk1.ibm.com"]
mailToList = []

hi = "<p class=MsoNormal><span style=\'font-size:12.0pt\'>Dear All,<o:p></o:p></span></p><p class=MsoNormal><span style=\'font-size:12.0pt\'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style=\'font-size:12.0pt\'>This is today's AOC incident number and SLA status.<o:p></o:p></span> </p>"
slmReminder = "<p class=MsoNormal><span style=\'font-size:12.0pt\'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style=\'font-size:12.0pt\'>{0} incident(s) will miss SLA the next working day. <o:p></o:p></span></p>"
updateReminder = "<p class=MsoNormal><span style=\'font-size:12.0pt\'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style=\'font-size:12.0pt\'>Please daily update incident in remedy. <o:p></o:p></span></p>"


def Main():
	logging.info("Generating AOC Open Inc Daily Update Check Report for {0}".format(today.strftime("%Y-%m-%d")))
	'''
	if (today.weekday() in (5, 6)):
		logging.info("Not business day. Program ends.")
		return
	'''
	if not DebugMode:
		time.sleep(10)
		filelist = [ f for f in os.listdir(workDir) if f.startswith("AOC Open Incident") and f.endswith(".xlsx") ]
		for f in filelist:
		    os.remove(workDir + f)
	

	outlook = Dispatch("Outlook.Application")
	mapi = outlook.GetNamespace("MAPI")
	inbox = mapi.GetDefaultFolder(const.olFolderInbox)
	remedy = inbox.Folders["Remedy"]

	incFullPath = r"C:\Users\IBM_ADMIN\Desktop\CPA-AOC\Daily Report INC 1600\AOC Open Incident - 05 Apr 2017.csv"
	incWDFullPath = r"C:\Users\IBM_ADMIN\Desktop\CPA-AOC\Daily Report INC 1600\AOC Open Incident Work Details - 05 Apr 2017.csv"

	for mail in remedy.Items:
		# Get the two emails sent at today
		rcvDate = PyTimetoDatetime(mail.ReceivedTime).date()
		
		logging.debug(mail.Subject + " @ " + rcvDate.strftime("%Y-%m-%d"))

		if rcvDate == date.today():
			if "AOC Open Incident Daily 16:00" == mail.Subject:
				mail.UnRead = False
				attachment = mail.Attachments.Item(1)
				incFullPath = workDir + attachment.FileName
				logging.info("incFullPath = " + incFullPath)
				attachment.SaveAsFile(incFullPath)
			elif "AOC Open Incident Work Details" == mail.Subject:
				mail.UnRead = False
				attachment = mail.Attachments.Item(1)
				incWDFullPath = workDir + attachment.FileName
				logging.info("incWDFullPath = " + incWDFullPath)
				attachment.SaveAsFile(incWDFullPath)

	# Excel Part
	excel = Dispatch("Excel.Application")

	wdWb = excel.Workbooks.Open(incWDFullPath)
	incWb = excel.Workbooks.Open(incFullPath)
	excel.DisplayAlerts = False
	excel.Visible = True if DebugMode else False
	excel.ScreenUpdating = True if DebugMode else False
	
	wdSheet = wdWb.ActiveSheet
	incSheet = incWb.ActiveSheet
	
	incIdCol1 = wdSheet.Columns(1) # Column A : Incident ID
	lastUpdateTimeCol1 = wdSheet.Columns(4) # Column D : Work Info Submit Time
	incIdCol2 = incSheet.Columns(1) # Column A : Incident ID
	assigneeCol2 = incSheet.Columns(6) # Column F : Assignee
	statusCol2 = incSheet.Columns(7) # Column G: Status
	lastUpdateTimeCol2 = incSheet.Columns(11) # Column K : Last Update Time
	slmCol2 = incSheet.Columns(8) # Column H : SLM Status

	# Convert PyTime to datetime.date
	lutDict = {} # last update time dict(str, datetime) to store (incId, lastUpdateTime)
	for i in range(2, incIdCol1.End(const.xlDown).Row + 1):
		lut1 = lastUpdateTimeCol1.Cells(i).Value
		lutDict[incIdCol1.Cells(i).Value] = PyTimetoDatetime(lut1)
		logging.debug(str(incIdCol1.Cells(i).Value) + " @ " + str(lastUpdateTimeCol1.Cells(i).Value))

	# Put last update time to wdSheet
	# List all not updated incidents
	notUpdatedIncList = []
	lastUpdateTimeCol2.Cells(1).Value = "Last Update Time"
	for i in range(2, incIdCol2.End(const.xlDown).Row + 1):
		incNo = str(incIdCol2.Cells(i).Value)
		lut2 = lutDict.get(incNo)
		
		if lut2 is not None:
			lastUpdateTimeCol2.Cells(i).Value = lut2.strftime("%Y-%m-%d %H:%M")
		if lut2 is None or lut2.date() < today:
			notUpdatedIncList.append(incNo)
			AddToRecipient(str(assigneeCol2.Cells(i).Value))
			logging.debug(str(incIdCol2.Cells(i).Value) +
				  str(assigneeCol2.Cells(i).Value) +
				  str(statusCol2.Cells(i).Value) +
				  str(lastUpdateTimeCol2.Cells(i).Value))
	logging.info("notUpdatedIncList = " + str(notUpdatedIncList))
	wdWb.Close(False)

	# Highlight warning & breached case
	logging.info("Highlighting SLM warning & breached case(s) begin.")
	slmBreachedList = []
	slmWarningList = []
	for slm in incSheet.Range(slmCol2.End(const.xlUp), slmCol2.End(const.xlDown)):
		if "Service Targets Breached" == str(slm.Value):
			incSheet.Range("A"+str(slm.Row)+":K"+str(slm.Row)).Interior.Color = RGBToInt(255, 0, 0)
			logging.info(incSheet.Range("A"+str(slm.Row)+":K"+str(slm.Row)).Value)
			slmBreachedList.append(str(incIdCol2.Cells(slm.Row).Value))
		if "Service Target Warning" == str(slm.Value):
			incSheet.Range("A"+str(slm.Row)+":K"+str(slm.Row)).Interior.Color = RGBToInt(255, 255, 0)
			AddToRecipient(str(assigneeCol2.Cells(slm.Row).Value))
			logging.info(incSheet.Range("A"+str(slm.Row)+":K"+str(slm.Row)).Value)
			slmWarningList.append(str(incIdCol2.Cells(slm.Row).Value))
	logging.info("slmBreachedList = " + str(slmBreachedList))
	logging.info("slmWarningList = " + str(slmWarningList))
	logging.info("Highlighting SLM warning & breached case(s) end.")
	
	# Statistics
	incSheet.Cells(1, 14).Value = "Total"
	incSheet.Cells(1, 15).Value = "Within SLA"
	incSheet.Cells(1, 16).Value = "SLA Warning"
	incSheet.Cells(1, 17).Value = "SLA Missed"
	incSheet.Cells(2, 13).Value = "CGO"
	incSheet.Cells(2, 14).Formula = '=SUM(O2:Q2)'
	incSheet.Cells(2, 15).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Cargo",$H:$H,"Within the Service Target")'
	incSheet.Cells(2, 16).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Cargo",$H:$H,"Service Target Warning")'
	incSheet.Cells(2, 17).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Cargo",$H:$H,"Service Targets Breached")'
	incSheet.Cells(3, 13).Value = "ENG"
	incSheet.Cells(3, 14).Formula = '=SUM(O3:Q3)'
	incSheet.Cells(3, 15).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Engineering",$H:$H,"Within the Service Target")'
	incSheet.Cells(3, 16).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Engineering",$H:$H,"Service Target Warning")'
	incSheet.Cells(3, 17).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Engineering",$H:$H,"Service Targets Breached")'
	incSheet.Cells(4, 13).Value = "FOP"
	incSheet.Cells(4, 14).Formula = '=SUM(O4:Q4)'
	incSheet.Cells(4, 15).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Flight Operations",$H:$H,"Within the Service Target")'
	incSheet.Cells(4, 16).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Flight Operations",$H:$H,"Service Target Warning")'
	incSheet.Cells(4, 17).Formula = '=COUNTIFS($E:$E,"IBM ASM-AOC Flight Operations",$H:$H,"Service Targets Breached")'

	# Adjust column width
	incSheet.Columns(1).ColumnWidth = 16
	incSheet.Columns(2).ColumnWidth = 15
	incSheet.Columns(3).ColumnWidth = 10
	incSheet.Columns(4).ColumnWidth = 5
	incSheet.Columns(5).ColumnWidth = 30
	incSheet.Columns(6).ColumnWidth = 12
	incSheet.Columns(7).ColumnWidth = 11
	incSheet.Columns(8).ColumnWidth = 23
	incSheet.Range("$I:$K").Columns.ColumnWidth = 16
	incSheet.Columns(12).ColumnWidth = 2
	incSheet.Columns(13).ColumnWidth = 5
	incSheet.Columns(14).ColumnWidth = 6
	incSheet.Columns(15).ColumnWidth = 10
	incSheet.Columns(16).ColumnWidth = 11
	incSheet.Columns(17).ColumnWidth = 11

	# Horizontailly Align center
	incSheet.Range("I:K").HorizontalAlignment = const.xlLeft
	incSheet.Range("M1:Q4").HorizontalAlignment = const.xlCenter

	# Add border
	for i in range(7,13):
		incSheet.Range("A1:K" + str(incIdCol2.End(const.xlDown).Row)).Borders(i).LineStyle = const.xlContinuous
		incSheet.Range("A1:K" + str(incIdCol2.End(const.xlDown).Row)).Borders(i).Weight = const.xlThin
		incSheet.Range("M1:Q4").Borders(i).LineStyle = const.xlContinuous
		incSheet.Range("M1:Q4").Borders(i).Weight = const.xlThin
	logging.info("Borders added.")
	
	# Fill color for header.
	incSheet.Range("A1:K1").Interior.Color = RGBToInt(0, 176, 240)
	incSheet.Range("M1:Q1").Interior.Color = RGBToInt(0, 176, 240)

	logging.info("Header color filled.")
	
	# Save as xlsx
	newFilename = incWb.FullName[:incWb.FullName.rfind(".")] + ".xlsx"
	incWb.SaveAs(newFilename, const.xlOpenXMLWorkbook)

	logging.info("Drafing an email...")
	# New an email
	newMail = outlook.CreateItem(const.olMailItem)
	newMail.To = ";".join(mailToList)
	newMail.CC = ";".join(mailCCList)
	# Convert incident number table to HTML
	newMail.Subject = "AOC Open Incidents Report " + today.strftime("%Y-%m-%d")
	
	newMail.Display()
	bodyHtml = Soup(newMail.HTMLBody, "html.parser")
	divHtml = bodyHtml.div
	insertPos = 0
	divHtml.insert(insertPos, Soup(hi, "html.parser"))
	insertPos += 1
	incNoHtml = Soup(RangetoHTML(incWb, incSheet, incSheet.Range("M1:Q4")), "html.parser")
	incSheet.Range("M1:Q4").Delete()
	bodyHtml.style.append(incNoHtml.style.string)
	divHtml.insert(insertPos, incNoHtml.table)
	insertPos += 1

	# Hide unuseful columns
	incSheet.Columns(10).Hidden = True
	incSheet.Columns(8).Hidden = True
	incSheet.Columns(5).Hidden = True
	incSheet.Columns(4).Hidden = True
	incSheet.Columns(3).Hidden = True
	incSheet.Columns(2).AutoFit()
	incSheet.Columns(6).AutoFit()

	# Post all SLM-warning incident
	slmWarningCount = len(slmWarningList)
	if len(slmWarningList) > 0:
		divHtml.insert(insertPos, Soup(slmReminder.format(slmWarningCount), "html.parser"))
		insertPos += 1
		for i in range(2, incIdCol2.End(const.xlDown).Row + 1):
			incNo = str(incIdCol2.Cells(i).Value)
			if incNo not in slmWarningList:
				incSheet.Rows(i).Hidden = True
			
		slmWarningHtml = Soup(RangetoHTML(incWb, incSheet, incSheet.UsedRange.SpecialCells(const.xlCellTypeVisible)), "html.parser")
		bodyHtml.style.append(slmWarningHtml.style.string)
		divHtml.insert(insertPos, slmWarningHtml.table)
		insertPos += 1
		# reset hidden
		incSheet.Rows.Hidden = False
	logging.info("SLW-warning incidents  added to HTML body.")
	
	# Post all not-updated incident
	if len(notUpdatedIncList) > 0:
		divHtml.insert(insertPos, Soup(updateReminder, "html.parser"))
		insertPos += 1
		for i in range(incIdCol2.End(const.xlDown).Row + 1, 1, -1):
			incNo = str(incIdCol2.Cells(i).Value)
			if incNo not in notUpdatedIncList:
				incSheet.Rows(i).Delete()

		notUpdatedHtml = Soup(RangetoHTML(incWb, incSheet, incSheet.UsedRange.SpecialCells(const.xlCellTypeVisible)), "html.parser")
		bodyHtml.style.append(notUpdatedHtml.style.string)
		divHtml.insert(insertPos, notUpdatedHtml.table)
		insertPos += 1
	logging.info("Not-updated incidents added to HTML body.")
	incWb.Close(False)
	
	newMail.HTMLBody = bodyHtml.decode()
	logging.info("Email body filled.")
	newMail.Attachments.Add(newFilename)
	logging.info("Attachment added.")
	# newMail.Send()
	logging.info("Email sent.")
	excel.Visible = 1
	# excel.Quit()
	del excel

	os.remove(incWDFullPath)
	os.remove(incFullPath)
	logging.info("inc.csv and incWd.csv removed.")

	logging.info("End")
	


def RGBToInt(r, g, b):
	'''
	Excel can use an integer calculated by the formula:
	Red + (Green * 256) + (Blue * 256 * 256)
	'''
	return r + (g * 256) + (b * 256 * 256)

def RangetoHTML(wb, ws, rng):
	''' Return the table element in html style. Can handle a multiple-area range. Keep original format. '''
	tmpSheet = wb.Worksheets.Add()
	rng.Copy(tmpSheet.Range("A1"))
	colCnt = 1
	for a in rng.Areas:
		for c in a.Columns:
			tmpSheet.Columns(colCnt).ColumnWidth = c.ColumnWidth
			colCnt += 1
	tmpFile = workDir + datetime.now().strftime("%Y-%m-%d %H-%M-%S-%f") + ".htm"
	p = wb.PublishObjects.Add(SourceType = const.xlSourceRange,
							  Filename = tmpFile,
							  Sheet = tmpSheet.Name,
							  Source = tmpSheet.UsedRange.Address,
							  HtmlType = const.xlHtmlStatic)
	p.Publish()
	html = open(tmpFile).read().replace("align=center\nx:publishsource=", "align=left\nx:publishsource=")
	tmpSheet.Delete()
	os.remove(tmpFile)
	return html
	
def PyTimetoDatetime(t):
	''' Convert PyTime (Date type used in Excel) into Python's datetime.datetime'''
	return datetime.strptime(t.Format("%Y-%m-%d %H:%M"), "%Y-%m-%d %H:%M")

def AddToRecipient(name):
	if name == "None":
		return
	if name == "Tommy Ho":
		name = "SPITYH"
	if name not in mailToList:
		mailToList.append(name)

logging.basicConfig(filename	= workDir + 'INC 1600.log',
					level		= logging.DEBUG,
					format		= "%(asctime)s %(name)s %(levelname)s\t - %(message)s")

	
if __name__ == "__main__":
	Main()
