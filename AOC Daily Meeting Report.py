from win32com.client import Dispatch, constants as const
from datetime import datetime, date, timedelta
import logging, os, time

DebugMode = False
today = date.today()
workDir = "C:\Users\IBM_ADMIN\Desktop\CPA-AOC\Daily Report Team Meeting\\"
mailToList = ["Lei Wang", "Ce Zheng", "Miao Zhao", "Cong Wang", "Xin Long He", "Hui Xia Tian", "Chu Jiang", "Jin Yu Yan"]

def Main():
    logging.info("Generating AOC Open Inc/SR Report for {0} begins.".format(today.strftime("%Y-%m-%d")))
    if (today.weekday() in (5, 6)):
        logging.info("Not business day. Program ends.")
        return
    if (not DebugMode):
        logging.info("Sleep 30s begins...")
        time.sleep(30)
        logging.info("Sleep 30s ends...")
    # Get attachments from outlook, save to workDir
    outlook = Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    remedy = mapi.GetDefaultFolder(const.olFolderInbox).Folders["Remedy"]

    logging.info("Get attachments from outlook")
    for mail in remedy.Items:
        rcvDate = PyTimeToDate(mail.ReceivedTime)
        if rcvDate == today:
            if "AOC Open Incident Daily 10:00" == mail.Subject:
                mail.UnRead = False
                attachment = mail.Attachments.Item(1)
                incFile = workDir + attachment.FileName
                attachment.SaveAsFile(incFile)
                logging.info("incFile = " + incFile)
            if "AOC Open SR Daily 10:00" == mail.Subject:
                mail.UnRead = False
                attachment = mail.Attachments.Item(1)
                srFile = workDir + attachment.FileName
                attachment.SaveAsFile(srFile)
                logging.info("srFile = " + srFile)

    newIncFile = incFile[:incFile.rfind(".")] + ".xlsx"
    newSrFile = srFile[:srFile.rfind(".")] + ".xlsx"
    
    # Cal business date and get the filename of last business day
    lastInc = "AOC Open Incident - " + str(LastBizDay(today).strftime("%d %b %Y"))
    lastSr = "AOC Open SR - " + str(LastBizDay(today).strftime("%d %b %Y"))
    # Open file of both days
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    
    incWb = excel.Workbooks.Open(incFile)
    incWs = incWb.ActiveSheet

    logging.info("Processing incident sheet " + str(incWs.Name))
    # Add header and vlookup formula into today's excel
    incRows = incWs.UsedRange.Rows.Count
    incEtaVlookup = '''=IFNA(VLOOKUP($A2,'{0}[{1}]{2}'!$A$1:$L$100,11,FALSE),"")'''.format(workDir, lastInc+".xlsx", lastInc)
    incUpdateVlookup = '''=IFNA(VLOOKUP($A2,'{0}[{1}]{2}'!$A$1:$L$100,12,FALSE),"")'''.format(workDir, lastInc+".xlsx", lastInc)
    incWs.Range("K1").Value = "Target Date"
    incWs.Range("K2").Formula = incEtaVlookup
    incWs.Range("K2").NumberFormat = "yyyy-mm-dd"
    incWs.Range("L1").Value = "Update"
    incWs.Range("L2").Formula = incUpdateVlookup
    incWs.Range("K2:L2").AutoFill(incWs.Range("K2:L"+str(incRows)))
    incWs.Range("K2:L"+str(incRows)).Copy()
    incWs.Range("K2:L"+str(incRows)).PasteSpecial(Paste = const.xlPasteValuesAndNumberFormats)
    incWs.Range("A1:L1").Interior.Color = RGBToInt(0, 176, 240)
    logging.info("Fill header and formula done.")
    # Adjust column width
    incWs.Columns(1).ColumnWidth = 16
    incWs.Columns(2).ColumnWidth = 10
    incWs.Columns(3).Hidden = True
    incWs.Columns(4).ColumnWidth = 5
    incWs.Range("E:F").Columns.AutoFit()
    incWs.Columns(8).Hidden = True
    incWs.Range("$I:$J").Columns.AutoFit()
    logging.info("Adjust column width done.")
    # Add border
    for i in range(7,13):
        incWs.UsedRange.Borders(i).LineStyle = const.xlContinuous
        incWs.UsedRange.Borders(i).Weight = const.xlThin
        incWs.UsedRange.Borders(i).LineStyle = const.xlContinuous
        incWs.UsedRange.Borders(i).Weight = const.xlThin
    logging.info("Borders added.")
    # Save as excel
    incWb.SaveAs(newIncFile, const.xlOpenXMLWorkbook)
    incWb.Close(True)
    logging.info("Incident process completed.")
    if not DebugMode:
        os.remove(workDir + lastInc + ".xlsx")
        logging.info("Last incident file deleted.")
    os.remove(incFile)
    logging.info("Incident.csv deleted.")
    
    
    srWb = excel.Workbooks.Open(srFile)
    srWs = srWb.ActiveSheet    
    srRows = srWs.UsedRange.Rows.Count
    logging.info("{0} SR(s) today".format(srRows-1))
    if srRows > 1:    
        srEtaVlookup = '''=IFNA(VLOOKUP($A2,'{0}[{1}]{2}'!$A$1:$L$100,10,FALSE),"")'''.format(workDir, lastSr+".xlsx", lastSr)
        srUpdateVlookup = '''=IFNA(VLOOKUP($A2,'{0}[{1}]{2}'!$A$1:$L$100,11,FALSE),"")'''.format(workDir, lastSr+".xlsx", lastSr)
        srWs.Range("J1").Value = "Target Date"
        srWs.Range("J2").Formula = srEtaVlookup
        srWs.Range("J2").NumberFormat = "yyyy-mm-dd"
        srWs.Range("K1").Value = "Update"
        srWs.Range("K2").Formula = srUpdateVlookup
        if srRows > 2:
            srWs.Range("J2:K2").AutoFill(srWs.Range("J2:K"+str(srRows)))
        srWs.Range("J2:K"+str(srRows)).Copy()
        srWs.Range("J2:K"+str(srRows)).PasteSpecial(Paste = const.xlPasteValuesAndNumberFormats)
        srWs.Range("A1:K1").Interior.Color = RGBToInt(0, 176, 240)

        srWs.Range("A:F").Columns.AutoFit()
        srWs.Range("G:I").Columns.Hidden = True

        for i in range(7,13):
            srWs.UsedRange.Borders(i).LineStyle = const.xlContinuous
            srWs.UsedRange.Borders(i).Weight = const.xlThin
            srWs.UsedRange.Borders(i).LineStyle = const.xlContinuous
            srWs.UsedRange.Borders(i).Weight = const.xlThin
    srWb.SaveAs(newSrFile, const.xlOpenXMLWorkbook)
    srWb.Close(True)
    logging.info("SR process completed.")
    # excel.Quit()
    del excel
    if not DebugMode:
        os.remove(workDir + lastSr + ".xlsx")
        logging.info("Last SR file deleted.")
    os.remove(srFile)
    logging.info("SR.csv deleted")
    
    
    # Send email
    logging.info("Draft an email")
    newMail = outlook.CreateItem(const.olMailItem)
    newMail.Display()
    newMail.Subject = "AOC Open INC/SR " + today.strftime("%d %b %Y")
    newMail.To = ";".join(mailToList)
    
    newMail.Attachments.Add(newIncFile)
    if srRows > 1:
        newMail.Attachments.Add(newSrFile)
        newMail.Body = "AOC Open INC/SR for daily meeting."
    else:
        newMail.Body = "AOC Open INC for daily meeting. No SR today."
    # newMail.Send()
    # logging.info("Email sent out")

    logging.info("Generating AOC Open Inc/SR Report for {0} ends.".format(today.strftime("%Y-%m-%d")))

def PyTimeToDate(t):
    ''' Convert PyTime (Date type used in Excel) into Python's datetime.datetime'''
    return datetime.strptime(t.Format("%Y-%m-%d %H:%M"), "%Y-%m-%d %H:%M").date()

def LastBizDay(day):
    shift = timedelta(days = max(1, (day.weekday() + 6) % 7 - 3))
    return day - shift

def RGBToInt(r, g, b):
    '''
    Excel can use an integer calculated by the formula:
    Red + (Green * 256) + (Blue * 256 * 256)
    '''
    return r + (g * 256) + (b * 256 * 256)

logging.basicConfig(filename	= workDir + 'Meeting Report 1000.log',
					level		= logging.INFO,
					format		= "%(asctime)s %(name)s %(levelname)s\t - %(message)s")

if __name__ == "__main__":
    Main()
