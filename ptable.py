import win32com.client as win32
win32c = win32.constants

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list):
    """ Create a pivot table.
        Parameters: 
            wb          Workbook reference.
            ws1         Source worksheet. 
            pt_ws       Target worksheet. 
            ws_name     Name of target worksheet. 
            pt_name     Name given to pivot table.
            pt_rows, pt_cols, pt_filters, pt_fields
                Values required to create the pivot table. """
    # Pivot table location.
    pt_loc = len(pt_filters) + 5
    
    # Source the pivot table data. 
    pc = wb.PivotCaches().Create( SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # Create the pivot table object.
    pc.CreatePivotTable( TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # Select the location to create the pivot table. 
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table. 
    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Assigns the values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity control Boolean. 
    pt_ws.PivotTables(pt_name).ShowValuesRow = False
    pt_ws.PivotTables(pt_name).ColumnGrand = False