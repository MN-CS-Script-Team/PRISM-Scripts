'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - CALI TO EXCEL.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED


'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

Dim CAFS_checkbox

BeginDialog CALI_to_excel_Dialog, 0, 0, 231, 115, "CALI To Excel"
  DropListBox 70, 15, 120, 15, "Run for your own CALI list"+chr(9)+"Run for another CALI list", action_dropdown
  CheckBox 15, 35, 200, 10, "Check here if you want to include last payment received", payments_checkbox
  CheckBox 15, 65, 115, 10, "Check here if you want to include total arrears, monthly accrual amount, and non-accrual amount to Excel", CAFS_checkbox
  ButtonGroup ButtonPressed
    OkButton 165, 70, 50, 15
    CancelButton 165, 90, 50, 15
  Text 25, 45, 120, 10, "(This takes more time to process)"
  Text 10, 15, 60, 10, "Select an action:"
  Text 30, 75, 120, 30, "total arrears, monthly accrual amount, and non-accrual amount to Excel from CAFS. "
  Text 30, 100, 120, 10, "(This takes more time to process)"
EndDialog

BeginDialog CALI_selection_dialog, 0, 0, 211, 80, "CALI Criteria"
  Text 5, 15, 205, 10, "Enter these fields to run this script on another CALI caseload:"
  Text 5, 35, 25, 10, "County:"
  EditBox 35, 30, 30, 15, cali_office
  Text 75, 35, 25, 10, "Team:"
  EditBox 105, 30, 25, 15, cali_team
  Text 145, 35, 30, 10, "Position:"
  EditBox 180, 30, 25, 15, cali_position
  ButtonGroup ButtonPressed
    OkButton 105, 55, 50, 15
    CancelButton 160, 55, 50, 15
EndDialog

'change team, position

'***********************************************************************************************************************************************
'If the user is already on the CALI screen when the script is run, results may be inaccurate.  Also, if the user runs the script when the 
'position listing screen is open, the screen must be exited before the script can run properly.  This function checks to see if either of 
'these circumstances apply.  If the position list is open, the script exits the list, and if the CALI screen is open, navigates away so that
'the report will function properly.
FUNCTION refresh_CALI_screen
	EMReadScreen check_for_position_list, 22, 8, 36
		IF check_for_position_list = "Caseload Position List" THEN
			PF3
		END IF
	EMReadScreen check_for_caseload_list, 13, 2, 32
		If check_for_caseload_list = "Caseload List" THEN	
			CALL navigate_to_PRISM_screen("MAIN")
			transmit
		END IF
END FUNCTION

'Connects to Bluezone
EMConnect ""

check_for_PRISM(TRUE)

DIALOG CALI_to_excel_Dialog
	IF ButtonPressed = 0 THEN StopScript
	
	IF action_dropdown = "Run for another CALI list" THEN
		Dialog CALI_selection_dialog
	END IF

	EMReadScreen check_for_position_list, 22, 8, 36
		IF check_for_position_list = "Caseload Position List" THEN
			PF3
		END IF
	EMReadScreen check_for_caseload_list, 13, 2, 32
		If check_for_caseload_list = "Caseload List" THEN	
			CALL navigate_to_PRISM_screen("MAIN")
			transmit
		END IF	
	CALL navigate_to_PRISM_screen("CALI")  'Navigate to CALI, remove any case number entered, and display the desired CALI listing
	EMWriteScreen "             ", 20, 58
	EMWriteScreen "  ", 20, 69
	EMWriteScreen CALI_office, 20, 18
	EMWriteScreen "001", 20, 30
	EMWriteScreen CALI_team, 20, 40
	EMWriteScreen CALI_position, 20, 49
	transmit

	EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
	error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
	IF error_message_on_bottom_of_screen <> "" THEN script_end_procedure("The caseload you entered is invalid.  The script will now end.")

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

ObjExcel.Cells(1, 1).Value = "Case Number"
ObjExcel.Cells(1, 2).Value = "Function"
ObjExcel.Cells(1, 3).Value = "Program"
ObjExcel.Cells(1, 4).Value = "Interstate?"
ObjExcel.Cells(1, 5).Value = "CP Name"
ObjExcel.Cells(1, 6).Value = "NCP Name"
If Payments_Checkbox = checked then ObjExcel.Cells(1, 7).Value = "Last Payment Date"
If CAFS_checkbox = checked then
	ObjExcel.Cells(1, 8).Value = "Amount Of Arrears"
	ObjExcel.Cells(1, 9).Value = "Monthly Accrual"
	ObjExcel.Cells(1, 10).Value = "Monthly Non-Accrual"
End If

'Autofitting columns
For col_to_autofit = 1 to 10
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'sets row to fill info into Excel
excel_row = 2
prism_row = 8
Do 'Loops script until the end of CALI
	'Copies Case Number, Function Type, Program Type, CP Name, and NCP Name to the Excel document
	EMReadScreen prism_case_number, 14, prism_row, 7 'Reads and copies case number
	EMReadScreen function_type, 2, prism_row, 23 'Reads and copies function type
	EMReadScreen program_type, 3, prism_row, 27 'Reads and copies program type
	EMReadScreen interstate_code, 1, prism_row, 33 'Reads and copies intersate code
	EMReadScreen CP_name, 26, prism_row, 38 'Reads and copies CP name
	pf11
	EMReadScreen NCP_name, 26, prism_row, 33 'Reads and copies NCP name
	pf10
	
	'Set rows in Excel for case number, funtion type, program type, CP name, and NCP name
	ObjExcel.Cells(excel_row, 1).Value = prism_case_number
	ObjExcel.Cells(excel_row, 2).Value = function_type
	ObjExcel.Cells(excel_row, 3).Value = program_type
	ObjExcel.Cells(excel_row, 4).Value = interstate_code
	ObjExcel.Cells(excel_row, 5).Value = CP_name
	ObjExcel.Cells(excel_row, 6).Value = NCP_name

	prism_row = prism_row + 1
	excel_row = excel_row + 1

	EmReadscreen end_of_data_check, 11, prism_row, 32
	If end_of_data_check = "End of Data" then exit do

	IF prism_row = 19 THEN 
		PF8
		prism_row = 8
	END IF
Loop Until end_of_data_check = "End of Data"

EMWriteScreen "PALC", 21, 18 
Transmit

excel_row = 2

If payments_checkbox = checked then
	Do
		prism_case_number = Trim(ObjExcel.Cells(excel_row, 1).Value)
		EMWriteScreen Left (prism_case_number, 10), 20, 9
		EMWriteScreen Right (prism_case_number, 2), 20, 20
		Transmit
		EMReadScreen last_payment_date, 8, 9, 59
		If last_payment_date = "        " then last_payment_date = "No Payments"
		ObjExcel.Cells(excel_row, 7).Value = last_payment_date
		excel_row = excel_row + 1

	Loop until prism_case_number = ""
End IF

excel_row = 2

If CAFS_checkbox = checked then

	EMWriteScreen "CAFS", 21, 18 
	Transmit

	excel_row = 2

	Do
		prism_case_number = Trim(ObjExcel.Cells(excel_row, 1).Value)
		EMWriteScreen Left (prism_case_number, 10), 4, 8
		EMWriteScreen Right (prism_case_number, 2), 4, 19
		EMWriteScreen "D", 3, 29
		Transmit 
		EMReadScreen amount_of_arrears, 10, 12, 68
		ObjExcel.Cells(excel_row, 8).Value = amount_of_arrears
		EMReadScreen monthly_accrual, 7, 9, 32
		ObjExcel.Cells(excel_row, 9).Value = monthly_accrual
		EMReadScreen monthly_non_accrual, 7, 10, 32
		ObjExcel.Cells(excel_row, 10).Value = monthly_non_accrual

		excel_row = excel_row + 1
	
	Loop until prism_case_number = ""
End If

script_end_procedure("Success!!")
