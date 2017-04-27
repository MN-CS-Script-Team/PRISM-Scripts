'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "list-generator---cali.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/22/2017", "The script has been updated to include double-checks so that the worker does not accidentally cancel the script. Additionally, the script has been updated to give the worker the ability to cancel the script after the second dialog.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

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
	cancel_confirmation

	IF action_dropdown = "Run for another CALI list" THEN
		Dialog CALI_selection_dialog
		cancel_confirmation
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
