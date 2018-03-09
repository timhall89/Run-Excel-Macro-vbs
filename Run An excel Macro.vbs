'*************************************************************
' Path to workbook (include trailing \).
Const WBPath = ""

' Name of workbook.
Const WBName = "Name.vbs"

' Name of Sub To run
Const SubName = "VBSCall"

' Boolean to keep workbook open after or not.
Const KeepOpen = False

' Password for write access (blank string if no path).
Const WritePassword = "D3v4t1cs"
'*************************************************************

RunExcelMacro WBPath, WBName, SubName, KeepOpen, WritePassword

'********************************************************************************
' Function to run an excel macro.
'********************************************************************************
Function RunExcelMacro(strWBPath, strWBName, strSubName, blnKeepOpen, strWritePW)
	
	On Error Resume Next
	
	Dim appXL, wb
	
	' Create new instance of XL.
	Set appXL = CreateObject("Excel.Application")
	
	' Open the workbook.
	If strWritePW = "" Then
		Set wb = appXL.Workbooks.Open(strWBPath & strWBName, False, False)
	Else
		Set wb = appXL.Workbooks.Open(strWBPath & strWBName, False, False,,,strWritePW)
	End If
	
	' Make xl Visible.
	appXL.Visible = True
	
	' Apply single quotes around wb name if it contains spaces.
	If InStr(1,WBName," ") > 0 Then strWBName = "'" & strWBName & "'"
	
	' Call the excel sub.
	appXL.Run strWBName & "!" & strSubName
	
	' If workook was not opened read only save the workbook.
	If Not wb.ReadOnly Then wb.Save
	
	' Close the workbook.
	wb.Close
	
	' Quite the instance of excel and set object to nothing.
	appXL.Quit
	Set appXL = Nothing
	
	' Return value for if there were any errors.
	RunExcelMacro = (Err = 0)
	
End Function
