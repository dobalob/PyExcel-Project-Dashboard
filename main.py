""" ***********************
* Excel Project Dashboard * 
*********************** """

""" Libraries: 
    win32com    Provides access to Windows APIs from Python. """
import os
import csv
import win32com.client as win32

def appendWorksheet( book, name): 
    """ Inserts a new worksheet before a given worksheet, or at the end of the 
        workbook. """
    sheet = book.Worksheets.Add( After=book.Sheets( book.Sheets.Count), Before=None)
    sheet.Name = name 
    return sheet

ABC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

def getColumnName( col): 
    """ Get the string name of an Excel column given its integer number. """
    name = []
    while col: 
        col, r = divmod( col-1, 26)
        name[:0] = ABC[r]
    return ''.join( name)

def main(): 
    """ """
    # Open Excel. 
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True

    # Read csv file. 
    fpath = os.path.dirname( os.path.abspath( __file__))
    _csvFilename = os.path.join( fpath, "project_data.csv")
    

    # Create a workbook. 
    print( r'%s' % _csvFilename)
    wb = excel.Workbooks.Open( _csvFilename)
    ws = wb.Worksheets(1)
    ws.Name = "Data"
 
    # Calculate end date from start date and duration. 
    lastRow = ws.UsedRange.Columns.Rows.Count
    nLastCol = ws.UsedRange.Columns.Count
    lastCol = getColumnName( nLastCol)
    
    Cols = []
    for i in range( nLastCol):
        if ws.Cells(1,i+1).Value=="Start Date":
            # Store column letter. 
            Cols.append( getColumnName(i+1))

        elif ws.Cells(1,i+1).Value=="Duration":
            # Store column letter. 
            Cols.append( getColumnName(i+1))
            Cols.append( getColumnName(i+2))

            # Insert new column.
            ws.Cells(1,i+1).Offset(1,2).EntireColumn.Insert()
            ws.Range("%s1" % Cols[2]).Value = "End Date"

            # Populate column.
            calcEndDate = "=TEXT(WORKDAY.INTL(Data!$%s2-1,Data!$%s2,1), \"dd/m/yyyy\")" % (Cols[0], Cols[1])
            ws.Range("%s2" % Cols[2]).Formula = calcEndDate
            ws.Range("%s2" % Cols[2]).AutoFill( ws.Range("%s2:%s%s" % (Cols[2], Cols[2], lastRow)))
            break

    
    # Calculate progress from duration and days completed. 
    nLastCol = ws.UsedRange.Columns.Count
    lastCol = getColumnName( nLastCol)
    
    Cols = []
    for i in range( nLastCol):
        if ws.Cells(1,i+1).Value=="Duration":
            # Store column letter. 
            Cols.append( getColumnName(i+1))

        elif ws.Cells(1,i+1).Value=="Days completed": 
            # Store column letter. 
            Cols.append( getColumnName(i+1))
            Cols.append( getColumnName(i+2))

            # Insert new column. 
            ws.Cells( 1,i+1).Offset(1,2).EntireColumn.Insert()
            ws.Range("%s1" % Cols[2]).Value = "Progress"

            # Populate column. 
            calcProgress = "=(Data!$%s2/Data!$%s2)" % (Cols[1], Cols[0])
            ws.Range("%s2" % Cols[2]).Formula = calcProgress
            ws.Range("%s2" % Cols[2]).AutoFill( ws.Range("%s2:%s%s" % (Cols[2], Cols[2], lastRow)))
            break

    # Format data.
    nLastCol = ws.UsedRange.Columns.Count
    lastCol = getColumnName( nLastCol)

    Cols = []
    for i in range( nLastCol): 
        if ws.Cells(1,i+1).Value=="Progress":
            # Store column letter. 
            Cols.append( getColumnName(i+1))

            # Format decimals to percentages.
            ws.Range("%s2" % Cols[-1]).NumberFormat = '0%'
            #ws.Cells(2,8).NumberFormat = '0%'
            ws.Range("%s2" % Cols[-1]).Copy()
            ws.Range("%s2:%s%s" % (Cols[-1], Cols[-1], lastRow)).PasteSpecial()
        
        if ws.Cells(1,i+1).Value=="Budget": 
            # Store column letter. 
            Cols.append( getColumnName(i+1))

            # Add separators to values. 
            ws.Range("%s2" % Cols[-1]).NumberFormat = '#,##0'
            ws.Range("%s2" % Cols[-1]).Copy()
            xlPasteFormats = -4122
            ws.Range("%s2:%s%s" % (Cols[-1], Cols[-1], lastRow)).PasteSpecial( Paste=xlPasteFormats)
        
    
        if ws.Cells(1,i+1).Value=="Actual": 
            # Store column letter. 
            Cols.append( getColumnName(i+1))

            # Add separators to values. 
            ws.Range("%s2" % Cols[-2]).Copy()
            xlPasteFormats = -4122
            ws.Range("%s2:%s%s" % (Cols[-1], Cols[-1], lastRow)).PasteSpecial( Paste=xlPasteFormats)
        
    
        if ws.Cells(1,i+1).Value=="Actual": 
            # Store column letter. 
            Cols.append( getColumnName(i+1))

            # Add separators to values. 
            ws.Range("%s2" % Cols[-2]).Copy()
            

    # Format headers. 
    #ws.Rows(1).Font.Bold = True
    ws.Range("A1:%s%s" % (lastCol, lastRow)).Columns.AutoFit()
   
    # Format data as table.
    ws.UsedRange.Select()
    #excel.Selection.Columns.Autofit()
    ws.ListObjects.Add().TableStyle = "TableStyleMedium2"
    
    # Create worksheet. 
    #ws = wb.Worksheets.Add()
    #ws.Name = "Dashboard"
    #
    # Format first worksheet. 
    #ws.cells(1,1).Value = "Project Management Dashboard"
    #ws.Cells(1,1).Font.Name = "Arial"
    #ws.Cells(1,1).Font.Size = 18
    #ws.Cells(1,1).Font.ColorIndex = 2
    #ws.Rows(1).RowHeight = 32
    #ws.Rows(1).VerticalAlignment = 2
    #ws.Rows(1).Interior.ColorIndex = 16

    # Save workbook to file. 
    _xlFilename = os.path.join( fpath, "project")
    wb.SaveAs( r'%s' % _xlFilename, FileFormat=51)

    #wb.Close()
    #excel.Application.Quit()

if __name__=="__main__": 
    main()





