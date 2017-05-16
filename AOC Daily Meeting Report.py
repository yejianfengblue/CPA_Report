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
    
    excel = Dispatch("Excel.Application")
    if not DebugMode:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
    else:
        excel.Visible = True
        excel.ScreenUpdating = True

    logging.info("Processing incident file " + str(incFile))
    incWb = excel.Workbooks.Open(incFile)
    incWs = incWb.ActiveSheet
    # Delete the first two blank lines
    incWs.Range("1:2").Delete()    
    # Set Font
    incWs.UsedRange.Font.Name = "Calibri"
    incWs.UsedRange.Font.Size = 11
    
    usedRange = incWs.Range(incWs.Cells(1,1), incWs.Columns(12).End(const.xlDown))
    # Adjust column width, wrap text, hide columns
    usedRange.WrapText = True
    incWs.Columns(1).ColumnWidth = 8
    incWs.Columns(2).ColumnWidth = 24
    incWs.Columns(3).ColumnWidth = 10
    incWs.Columns(4).ColumnWidth = 4
    incWs.Columns(5).ColumnWidth = 6
    incWs.Columns(6).AutoFit()
    incWs.Columns(7).Hidden = True
    incWs.Columns(8).Hidden = True
    incWs.Columns(9).AutoFit()
    incWs.Columns(10).Hidden = True
    incWs.Columns(11).AutoFit()
    incWs.Columns(12).ColumnWidth = 40
    usedRange.Rows.AutoFit()
    usedRange.VerticalAlignment = const.xlCenter
    logging.info("Adjust column width done.")
    # Add border
    for i in range(7,13):
        usedRange.Borders(i).LineStyle = const.xlContinuous
        usedRange.Borders(i).Weight = const.xlThin
        usedRange.Borders(i).LineStyle = const.xlContinuous
        usedRange.Borders(i).Weight = const.xlThin
    logging.info("Borders added.")
    # Save as excel
    incWb.SaveAs(newIncFile, const.xlOpenXMLWorkbook)
    incWb.Close(True)
    logging.info("Incident process completed.")
    os.remove(incFile)
    logging.info("Incident.xls deleted.")


    logging.info("Processing SR file " + str(srFile))
    srWb = excel.Workbooks.Open(srFile)
    srWs = srWb.ActiveSheet
    srRows = srWs.UsedRange.Rows.Count
    logging.info("{0} SR(s) today".format(srRows-1))
    if srRows > 1: 
        # Delete the first two blank lines
        srWs.Range("1:2").Delete()    
        # Set Font
        srWs.UsedRange.Font.Name = "Calibri"
        srWs.UsedRange.Font.Size = 11
        
        usedRange = srWs.Range(srWs.Cells(1,1), srWs.Columns(10).End(const.xlDown))
        # Adjust column width, wrap text, hide columns
        usedRange.WrapText = True
        srWs.Columns(1).ColumnWidth = 8
        srWs.Columns(2).ColumnWidth = 24
        srWs.Columns(3).AutoFit()
        srWs.Columns(4).AutoFit()
        srWs.Columns(5).Hidden = True
        srWs.Columns(6).Hidden = True
        srWs.Columns(7).AutoFit()
        srWs.Columns(8).Hidden = True
        srWs.Columns(9).ColumnWidth = 10
        srWs.Columns(10).ColumnWidth = 50
        usedRange.Rows.AutoFit()
        usedRange.VerticalAlignment = const.xlCenter
        logging.info("Adjust column width done.")
        # Add border
        for i in range(7,13):
            usedRange.Borders(i).LineStyle = const.xlContinuous
            usedRange.Borders(i).Weight = const.xlThin
            usedRange.Borders(i).LineStyle = const.xlContinuous
            usedRange.Borders(i).Weight = const.xlThin
        logging.info("Borders added.")
    # Save as excel
    srWb.SaveAs(newSrFile, const.xlOpenXMLWorkbook)
    srWb.Close(True)
    logging.info("SR process completed.")
    os.remove(srFile)
    logging.info("SR.xls deleted.")
    excel.Visible = 1
    del excel
 
    
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
