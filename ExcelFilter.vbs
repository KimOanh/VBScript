private Function ExcelFilter(ByVal strPath, ByVal strCell, ByVal arrCondition, ByVal intFilterType)
	On Error Resume Next

	Set objExl = CreateObject("Excel.Application")
	objExl.Application.Visible = true

	Set wbBook = objExl.Workbooks.Open(strPath)
	Set wshSheet = wbBook.Worksheets(1)

	If Err.Number <> 0 then
	WbScript.Echo Err.Description
	Else
	wshSheet.Range(strCell).AutoFilter 1, arrCondition, intFilterType

	wbBook.Save
	wbBook.Close
	objExl.Quit
	End If

End Function