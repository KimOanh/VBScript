private Function FindCell(ByVal strPath, ByVal strSheet, ByVal strRange, ByVal strKeySearch)
'Description: Find a cell address in a Excel file by the value
'
'Input:
'
'	@strPath: Excel file path 
'	@strSheet: The sheet you want to find
'	@strRange: Search range
'	@strKeySearch: Search value
'
'Output:
'
'	@strAdrr: The cell address that store search result

	On Error Resume Next

	Dim strAdrr, findResult
	set objExl = CreateObject("Excel.Application")
	objExl.Application.Visible = true

	Set wbBook = objExl.Workbooks.Open(strPath)
	Set wshSheet = wbBook.Worksheets(strSheet)
	WbScript.Echo strRange

	Set findResult = wshSheet.Range(strRange).Find(strKeySearch)

	strAdrr = cStr(Replace(wshSheet.Cells(findResult.Row, findResult.Column).Address,"$",""))

	wbBook.Close
	objExl.Quit

	FindCell = strAdrr
End Function