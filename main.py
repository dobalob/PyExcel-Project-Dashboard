""" ***********************
* Excel Project Dashboard * 
*********************** """

""" Libraries: 
    win32com    Provides access to Windows APIs from Python. """
import win32com.client as win32

def createWorksheet( book, name): 
    """ Inserts a new worksheet before a given worksheet, or at the end of the 
        workbook. """
    sheet = book.Worksheets.Add( After=book.Sheets( book.Sheets.Count), Before=None)
    sheet.Name = name 
    return sheet

def main(): 
    """ """
    # Open Excel. 
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True

    # Create a workbook. 
    wb = excel.Workbooks.Add()

    # Create first worksheet. 
    ws = createWorksheet( wb, "Dashboard")
    for i in range( 3):
        wb.Worksheets(1).Delete()

    # Format first worksheet. 
    ws.cells(1,1).Value = "Project Management Dashboard"
    ws.Cells(1,1).Font.Name = "Arial"
    ws.Cells(1,1).Font.Size = 18
    ws.Cells(1,1).Font.ColorIndex = 2
    ws.Rows(1).RowHeight = 32
    ws.Rows(1).VerticalAlignment = 2
    ws.Rows(1).Interior.ColorIndex = 16


if __name__=="__main__": 
    main()





