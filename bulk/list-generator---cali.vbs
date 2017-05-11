'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "list-generator---cali.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSE											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		END IF
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			SET fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
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
		SET run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		SET fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
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
CALL changelog_update("04/27/2017", "The dialog box has been enhanced, the spreadsheet's formatting & readability have been improved, several bugs have been fixed, and the script will no longer end if the user inputs invalid caseload info.", "Kyle Nelson, Hennepin County")
CALL changelog_update("02/22/2017", "The script has been updated to include double-checks so that the worker does not accidentally cancel the script. Additionally, the script has been updated to give the worker the ability to cancel the script after the second dialog.", "Robert Fewins-Kalb, Anoka County")
CALL changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog CALI_to_excel_Dialog, 0, 0, 186, 115, "CALI To Excel"
  DropListBox 10, 15, 100, 15, "My caseload"+chr(9)+"Someone else's caseload", action_dropdown
  CheckBox 10, 45, 115, 10, "Date of last payment from PALC", payments_checkbox
  CheckBox 10, 55, 90, 10, "Financial info from CAFS", CAFS_checkbox
  ButtonGroup ButtonPressed
    OkButton 70, 95, 50, 15
    CancelButton 125, 95, 50, 15
  Text 10, 5, 100, 10, "Create a caseload report for:"
  GroupBox 5, 35, 175, 55, "Additional report options (check all that apply):"
  Text 10, 75, 165, 10, "Note: Additional options increase processing time"
EndDialog

BeginDialog CALI_selection_dialog, 0, 0, 211, 80, "CALI Criteria"
  EditBox 35, 30, 30, 15, cali_office
  EditBox 105, 30, 25, 15, cali_team
  EditBox 180, 30, 25, 15, cali_position
  ButtonGroup ButtonPressed
    OkButton 105, 55, 50, 15
    CancelButton 160, 55, 50, 15
  Text 5, 15, 205, 10, "Enter these fields to run this script on another CALI caseload:"
  Text 5, 35, 25, 10, "County:"
  Text 75, 35, 25, 10, "Team:"
  Text 145, 35, 30, 10, "Position:"
EndDialog

'***********************************************************************************************************************************************

'Connects to Bluezone
EMConnect ""

check_for_PRISM(TRUE)

'If you navigate from CALI to CALI, the caseload data on CALI doesn't refresh; if you navigate to CALI from any other screen, the caseload data refreshes to the user's caseload data
'After the initial dialog box is run, CALI needs to display the user's caseload data, so here the script makes sure to navigate away from the Caseload Position List (a CALI submenu) and CALI
EMReadScreen check_for_position_list, 22, 8, 36
IF check_for_position_list = "Caseload Position List" THEN PF3
EMReadScreen check_for_caseload_list, 13, 2, 32
IF check_for_caseload_list = "Caseload List" THEN
	CALL navigate_to_PRISM_screen("MAIN")
	transmit
END IF

'Run the initial dialog
DIALOG CALI_to_excel_Dialog
cancel_confirmation

'Navigate back to CALI, ensuring that the default caseload data on CALI is the user's own caseload data
CALL navigate_to_PRISM_screen("CALI")

'If user is running the report for someone else's caseload, run the second dialog box
IF action_dropdown = "Someone else's caseload" THEN
	Dialog CALI_selection_dialog
	cancel_confirmation
END IF

'Navigate to CALI, clear the case number field, and transmit the caseload data into CALI, and loops until valid caseload data is entered
DO
	CALL navigate_to_PRISM_screen("CALI")
	EMWriteScreen "             ", 20, 58
	EMWriteScreen "  ", 20, 69
	EMWriteScreen "001", 20, 30
	EMWriteScreen CALI_office, 20, 18
	EMWriteScreen CALI_team, 20, 40
	EMWriteScreen CALI_position, 20, 49
	transmit
	EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
	error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
	IF error_message_on_bottom_of_screen <> "" THEN
		MsgBox "Please enter a valid caseload."
		Dialog CALI_selection_dialog
		cancel_confirmation
	END IF
LOOP UNTIL error_message_on_bottom_of_screen = ""		

'Create an Excel spreadsheet for the script to write PRISM info to
SET ObjExcel = CreateObject("Excel.Application")
ObjExcel.Visible = True 																												'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
SET ObjWorkbook = objExcel.Workbooks.Add()
ObjExcel.DisplayAlerts = True 																									'Set this to false to make alerts go away. This is necessary in production.

'Write row 1 headers into Excel spreasheet'
ObjExcel.Cells(1, 1).Value = "Case Number"
ObjExcel.Cells(1, 2).Value = "Function"
ObjExcel.Cells(1, 3).Value = "Program"
ObjExcel.Cells(1, 4).Value = "Interstate?"
ObjExcel.Cells(1, 5).Value = "CP Name"
ObjExcel.Cells(1, 6).Value = "NCP Name"
IF Payments_Checkbox = checked THEN ObjExcel.Cells(1, 7).Value = "Last Payment Date"
IF CAFS_checkbox = checked THEN
	ObjExcel.Cells(1, 8).Value = "Amount Of Arrears"
	ObjExcel.Cells(1, 9).Value = "Monthly Accrual"
	ObjExcel.Cells(1, 10).Value = "Monthly Non-Accrual"
END IF

'Set values to excel_row and prism_row, for reading data from PRISM & writing to Excel
excel_row = 2
prism_row = 8

'Read info from CALI screen and add it to Excel
DO
	EMReadScreen prism_case_number_left, 10, prism_row, 7 												'Reads and copies first 10 digits of case number
	EMReadScreen prism_case_number_right, 2, prism_row, 19 												'Reads and copies last 2 digits of case number'
	prism_case_number = prism_case_number_left & " " & prism_case_number_right 		'Sets case number with only one space between the first 10 digits and the last 2 digits of the case number
	EMReadScreen function_type, 2, prism_row, 23
	EMReadScreen program_type, 3, prism_row, 27
	EMReadScreen interstate_code, 1, prism_row, 33
	EMReadScreen CP_name, 26, prism_row, 38
	pf11
	EMReadScreen NCP_name, 26, prism_row, 33
	pf10

	'Write data from CALI into Excel
	ObjExcel.Cells(excel_row, 1).Value = prism_case_number
	ObjExcel.Cells(excel_row, 2).Value = function_type
	ObjExcel.Cells(excel_row, 3).Value = program_type
	ObjExcel.Cells(excel_row, 4).Value = interstate_code
	ObjExcel.Cells(excel_row, 5).Value = CP_name
	ObjExcel.Cells(excel_row, 6).Value = NCP_name

	'Increment row values for the next loop'
	prism_row = prism_row + 1
	excel_row = excel_row + 1

	'Check to see if the script has reached the end of CALI
	EmReadscreen end_of_data_check, 11, prism_row, 32
	IF end_of_data_check = "End of Data" THEN EXIT DO

	'Check to see if the script has reached the end of a page on CALI. If so, script will PF8 and reset prism_row back to 8
	IF prism_row = 19 THEN
		PF8
		prism_row = 8
	END IF
LOOP UNTIL end_of_data_check = "End of Data"

'If user requests the date of last payment, then navigate to PALC, read that data and add it to Excel spreadsheet
IF payments_checkbox = checked THEN
	navigate_to_PRISM_screen("PALC")
	excel_row = 2
	DO
		prism_case_number = Trim(ObjExcel.Cells(excel_row, 1).Value)
		IF prism_case_number = "" THEN EXIT DO																			'Exit the do loop here to prevent an extra line from being written to the end of Excel spreadsheet
		EMWriteScreen Left (prism_case_number, 10), 20, 9
		EMWriteScreen Right (prism_case_number, 2), 20, 20
		Transmit
		EMReadScreen last_payment_date, 8, 9, 59
		IF last_payment_date = "        " THEN last_payment_date = "No Payments"
		ObjExcel.Cells(excel_row, 7).Value = last_payment_date
		ObjExcel.Cells(excel_row, 7).HorizontalAlignment = -4152										'Formats the cell so that all cells in the row are right-aligned
		excel_row = excel_row + 1
	LOOP UNTIL prism_case_number = ""
END IF

'Add financial data from CAFS if the user checked that option in the initial dialog box
IF CAFS_checkbox = checked THEN
	navigate_to_PRISM_screen("CAFS")
	excel_row = 2
	DO
		prism_case_number = Trim(ObjExcel.Cells(excel_row, 1).Value)
		IF prism_case_number = "" THEN EXIT DO																			'Exit the do loop here to prevent an extra line from being written to the end of Excel spreadsheet
		EMWriteScreen Left (prism_case_number, 10), 4, 8
		EMWriteScreen Right (prism_case_number, 2), 4, 19
		EMWriteScreen "D", 3, 29
		Transmit
		EMReadScreen amount_of_arrears, 10, 12, 68
		ObjExcel.Cells(excel_row, 8).Value = amount_of_arrears
		ObjExcel.Cells(excel_row, 8).NumberFormat = "0.00"													'Formats the cell so that all cells in the row are numbers with two decimal places'
		EMReadScreen monthly_accrual, 7, 9, 32
		ObjExcel.Cells(excel_row, 9).Value = monthly_accrual
		ObjExcel.Cells(excel_row, 9).NumberFormat = "0.00"													'Formats the cell so that all cells in the row are numbers with two decimal places'
		EMReadScreen monthly_non_accrual, 7, 10, 32
		ObjExcel.Cells(excel_row, 10).Value = monthly_non_accrual
		ObjExcel.Cells(excel_row, 10).NumberFormat = "0.00"													'Formats the cell so that all cells in the row are numbers with two decimal places'
		excel_row = excel_row + 1
	LOOP UNTIL prism_case_number = ""
END IF

'Read the data in row 1, and if a cell is empty, delete the column. Since data will always be added to columns 1-6 (A-F), reads from column 15 to 7 (O-G).
FOR excel_column = 15 TO 7 STEP -1
	check_for_data = Trim(ObjExcel.Cells(1, excel_column).Value)
	IF check_for_data = "" THEN ObjExcel.columns(excel_column).EntireColumn.Delete
NEXT

'Autofit columns so that the columns will be appropriately sized after all the data has been added to the spreadsheet
FOR col_to_autofit = 1 TO 10
	ObjExcel.columns(col_to_autofit).AutoFit()
NEXT

script_end_procedure("Success!!")