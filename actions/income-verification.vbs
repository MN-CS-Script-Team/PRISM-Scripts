'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "income-verification.vbs"
start_time = timer


'LOADING ROUTINE FUNCTIONS (FOR PRISM)--- UPDATED 9/8/16 to MASTER FUNCLIB--------------------------------------------------------------
IF IsEmpty(FuncLib_URL) = TRUE THEN 'Shouldn't load FuncLib if it already loaded once
    IF run_locally = FALSE or run_locally = "" THEN    'If the scripts are set to run locally, it skips this and uses an FSO below.
        IF use_master_branch = TRUE THEN               'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        Else                                            'Everyone else should use the release branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        End if
        SET req = CreateObject("Msxml2.XMLHttp.6.0")                'Creates an object to get a FuncLib_URL
        req.open "GET", FuncLib_URL, FALSE                          'Attempts to open the FuncLib_URL
        req.send                                                    'Sends request
        IF req.Status = 200 THEN                                    '200 means great success
            Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
            Execute req.responseText                                'Executes the script code
        ELSE                                                        'Error message
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
				
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/02/2016", "The script has been updated to improve the user experience by providing a more consistent error message handling and by ensuring the script is always writing the date on DDPL and PALC in the correct format.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("12/01/2016", "Dialog and Write fixes.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'Pulling in phone number

EMConnect ""
CALL check_for_PRISM(True)
Call navigate_to_PRISM_screen("CAAD")
EMsetcursor 3, 53
PF1
EMReadScreen Worker_phone_number, 12, 8, 35
Transmit

Date_complete = Date & ""
Worker_title = "Child Support Officer"

'dialog box to select the information needed

BeginDialog child_support_income_verification, 0, 0, 241, 265, "Child Support Income Verification"
  Text 10, 10, 50, 10, "Case Number"
  EditBox 60, 5, 145, 15, PRISM_case_number
  Text 15, 30, 105, 10, "Number of Months of Payments"
  CheckBox 45, 40, 50, 15, "3 months", three_months_checkbox
  CheckBox 45, 60, 60, 10, "6 months", six_months_checkbox
  CheckBox 45, 75, 55, 15, "12 months", twelve_months_checkbox
  Text 20, 95, 70, 10, "Custom Date Range"
  EditBox 25, 115, 65, 15, begin_date
  EditBox 110, 115, 70, 15, End_date
  Text 20, 150, 50, 10, "Date complete"
  EditBox 15, 165, 85, 15, Date_complete
  Text 125, 150, 95, 10, "Worker's Signature"
  EditBox 120, 165, 110, 15, worker_signature
  Text 15, 190, 80, 10, "Worker's Phone number"
  EditBox 15, 205, 95, 15, Worker_phone_number
  Text 120, 190, 55, 10, "Worker's Title"
  EditBox 120, 205, 110, 15, Worker_title
  ButtonGroup ButtonPressed
    OKButton 125, 235, 50, 15
    CancelButton 180, 235, 50, 15
  Text 95, 120, 10, 10, "to"
EndDialog


'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

call PRISM_case_number_finder(PRISM_case_number)

'Case number display dialog
Do
	'err_msg handling
	err_msg = ""
	Dialog child_support_income_verification
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then err_msg = err_msg & vbNewLine & "* Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	IF begin_date <> "" AND IsDate(begin_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* You entered a value for the beginning of the date range, but the script does not recognize it as a valid date."
	IF end_date <> "" AND IsDate(end_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* You entered a value for the end of the date range, but the script does not recognize it as a valid date."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."
Loop until err_msg = ""


'collecting information for the word document
'CP Name
call navigate_to_PRISM_screen("CPDE")
EMReadScreen CP_F, 12, 8, 34
EMReadScreen CP_M, 12, 8, 56
EMReadScreen CP_L, 17, 8, 8
EMReadScreen CP_S, 3, 8, 74
client_name = fix_read_data(CP_F) & " " & fix_read_data(CP_M) & " " & fix_read_data(CP_L)
If trim(CP_S) <> "" then client_name = client_name & " " & ucase(fix_read_data(CP_S))
client_name = trim(client_name)

'CP Address
'Navigating to CPDD to pull address info
call navigate_to_PRISM_screen("CPDD")
EMReadScreen address_line1, 30, 15, 11
EMReadScreen address_line2, 30, 16, 11
EMReadScreen city_state_zip, 49, 17, 11

'Cleaning up address info
address_line1 = replace(address_line1, "_", "")
call fix_case(address_line1, 1)
address_line2 = replace(address_line2, "_", "")
if trim (address_line2) <> "" then
	address = address_line1 & chr(13) & address_line2
else
	address = address_line1
end if
call fix_case(address_line2, 1)
city_state_zip = replace(replace(replace(city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(city_state_zip, 2)

'Monthly Accrual on the case
'Navigating to CAFS
call navigate_to_PRISM_screen("CAFS")
EMReadScreen monthly_accrual, 14, 9, 25
EMReadScreen monthly_nonaccrual, 14, 10, 25

monthly_accrual = FormatCurrency(monthly_accrual)
monthly_nonaccrual = FormatCurrency(monthly_nonaccrual)


CALL navigate_to_PRISM_screen("DDPL")


IF begin_date <> "" THEN CALL write_date(begin_date, "MM/DD/YYYY", 20, 38)
IF end_date <> "" THEN CALL write_date(end_date, "MM/DD/YYYY", 20, 67)
IF three_months_checkbox = checked THEN CALL write_date(DateAdd("m", -3, date), "MM/DD/YYYY", 20, 38)    '*****VERONICA THIS IS WHERE WE WOULD NEED THE CODE FOR THE 30, 60, 90 DAY CHECK BOX INFORMATION
IF six_months_checkbox = checked THEN CALL write_date(DateAdd("m", -6, date), "MM/DD/YYYY", 20, 38)
IF twelve_months_checkbox = checked THEN CALL write_date(DateAdd("m", -12, date), "MM/DD/YYYY", 20, 38)

transmit

row = 8
total_amount_issued = 0

Do
	EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have
	EMReadScreen direct_deposit_issued_date, 9, row, 11 				'Reading the issue date
	EMReadScreen direct_deposit_amount, 10, row, 33 				'Reading amount issued

	total_amount_issued = abs(total_amount_issued + abs(direct_deposit_amount)) 	'Totals amount issued

	row = row + 1 										'Increases the row variable by one, to check the next row

	EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have

	If row = 19 then 									'Resets row and PF8s
		PF8
		row = 8
	End if
Loop until end_of_data_check = "*** End of Data ***"

PF9												'Print DDPL for the time period

Transmit

total_amount_issued = FormatCurrency(total_amount_issued)

CALL navigate_to_PRISM_screen("PALC")

IF begin_date <> "" THEN CALL write_date(begin_date, "MM/DD/YYYY", 20, 35)
IF end_date <> "" THEN CALL write_date(end_date, "MM/DD/YYYY", 20, 49)
IF three_months_checkbox = checked THEN CALL write_date(DateAdd("m", -3, date), "MM/DD/YYYY", 20, 35)    		'*****VERONICA THIS IS WHERE WE WOULD NEED THE CODE FOR THE 30, 60, 90 DAY CHECK BOX INFORMATION
IF six_months_checkbox = checked THEN CALL write_date(DateAdd("m", -6, date), "MM/DD/YYYY", 20, 35)
IF twelve_months_checkbox = checked THEN CALL write_date(DateAdd("m", -12, date), "MM/DD/YYYY", 20, 35)

transmit

PALC_row = 9
case_total = 0
Do
	EMReadScreen end_of_data_check, 19, PALC_row, 28 					'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do 			'Exits do if we have
	EMReadScreen case_alloc_amt, 9, PALC_row, 70 							'Reading the payment IDEMReadScreen case_alloc_amt, 10, row, 70 					   	'Reading amount issued

	case_total = abs(case_total + abs(case_alloc_amt)) 					'Case Totals amount issued

	PALC_row = PALC_row + 1 									'Increases the row variable by one, to check the next row

	EMReadScreen end_of_data_check, 19, PALC_row, 28 					'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do 			'Exits do if we have

	If PALC_row = 19 then 									'Resets row and PF8s
		PF8
		PALC_row = 10
	End if
Loop until end_of_data_check = "*** End of Data ***"

case_total = FormatCurrency(case_total)

'Child's Name
Prism_row = 8

call navigate_to_PRISM_screen("CHDE")
EMWriteScreen "B", 3, 29
Transmit

Do
	EMReadScreen MCI_List2, 10, Prism_row, 67
	IF MCI_List2 = "          " Then exit Do
	Prism_row = Prism_row + 1
	IF MCI_List1 = "" THEN
		MCI_List1 = MCI_List2
	ELSE
		MCI_List1 = MCI_List1 & "," & MCI_List2
	END IF
Loop until MCI_List2 = "          "

MCIArray = split(MCI_List1, ",")

call navigate_to_PRISM_screen("CHDE")
For each MCI in MCIArray
	EMWriteScreen "D", 3, 29
	EMWriteScreen MCI, 4, 7
	Transmit
	EMReadScreen Child_F, 12, 9, 34
	EMReadScreen Child_M, 12, 9, 56
	EMReadScreen Child_L, 17, 9, 8
	EMReadScreen Child_S, 3, 9, 74
	child_name = fix_read_data(Child_F) & " " & fix_read_data(Child_M) & " " & fix_read_data(Child_L)
	If trim(Child_S) <> "" then child_name = child_name & " " & ucase(fix_read_data(Child_S))
	child_name = trim(child_name)
	child_list = child_list & child_name & VBnewline
Next

'MCI Array 0 = MCINumber
'MCI Array 1 = Child_F
'MCI Array 2 = Child_M
'MCI Array 3 = Child_L
'MCI Array 4 = Child_S

'Only need the next two lines once (opens word)
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

'repeat this for each document (opens the document)
Set objDoc = objWord.Documents.Add(word_documents_folder_path & "child-support-income-verification-form.docx")
With objDoc
	.FormFields("client_name").Result = client_name
	.FormFields ("address_line1").Result = address_line1
	.FormFields("city_state_zip").Result = city_state_zip
	.FormFields("total_amount_issued").Result = total_amount_issued
	.FormFields("case_total").Result = case_total
	.FormFields("PRISM_case_number").Result = PRISM_case_number
	.FormFields("monthly_accrual").Result = monthly_accrual
	.FormFields("monthly_nonaccrual").Result = monthly_nonaccrual
	.FormFields("child_name").Result = child_list
	.FormFields("worker_signature").Result = worker_signature
	.FormFields("Worker_title").Result = Worker_title
	.FormFields("Phone_number_editbox").Result = Worker_phone_number
	.FormFields("Date_complete").Result = Date_complete
	'Ect to fill in all the blanks in the documents
End With


script_end_procedure("")
