import win32com.client

office = win32com.client.Dispatch("Excel.Application")
wb = office.Workbooks.Open(r"SR.xlsx")

count = wb.Sheets.Count
for i in range(count):
	ws = wb.Worksheets[i]
	ws.Unprotect() # IF protected

	pivotCount = ws.PivotTables().Count
	for j in range(1, pivotCount+1):
		ws.PivotTables(j).PivotCache().Refresh()

    # Put protection back on
	ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, AllowUsingPivotTables=True)

	ws.PrintOut()
	print "Worksheet: %s - has been sent to the printer" % (ws.Name)