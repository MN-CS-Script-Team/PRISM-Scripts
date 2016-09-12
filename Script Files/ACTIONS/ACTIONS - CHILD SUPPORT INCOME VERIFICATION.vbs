'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CHILD SUPPORT INCOME VERIFICATION.vbs"
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

'dialog box to select the information needed

BeginDialog Child_Support_verif_dialog, 0, 0, 241, 265, "Child Support Verif Dialog"
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
  EditBox 15, 205, 95, 15, Phone_number_editbox
  Text 120, 190, 55, 10, "Worker's Title"
  EditBox 120, 205, 110, 15, Worker_title
  ButtonGroup ButtonPressed
    PushButton 125, 235, 50, 15, "OK", OK
    PushButton 180, 235, 50, 15, "Cancel", Cancel
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
	Dialog child_support_income_verification
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
Loop until case_number_valid = True


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

CALL navigate_to_PRISM_screen("DDPL")


IF begin_date <> "" THEN EMWriteScreen begin_date, 20, 038
IF end_date <> "" THEN EMWriteScreen end_date, 20, 067
IF three_months_checkbox = checked THEN EMWriteScreen DateAdd("m", -3, date), 20, 038 
IF six_months_checkbox = checked THEN EMWriteScreen DateAdd("m", -6, date), 20, 038
IF twelve_months_checkbox = checked THEN EMWriteScreen DateAdd("m", -12, date), 20, 038

transmit

row = 8

total_amount_issued = 0

Do

EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have
EMReadScreen direct_deposit_issued_date, 9, row, 11 				'Reading the issue date
EMReadScreen direct_deposit_amount, 10, row, 33 				'Reading amount issued

total_amount_issued = FormatCurrency(total_amount_issued + abs(direct_deposit_amount)) 	'Totals amount issued

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


'Only need the next two lines once (opens word)
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

'repeat this for each document (opens the document)
Set objDoc = objWord.Documents.Add(word_documents_folder_path & "Child Support Income Verification Form.docx")
With objDoc
		.FormFields("client_name").Result = client_name
		.FormFields ("address_line1").Result = address_line1
		.FormFields("city_state_zip").Result = city_state_zip
		.FormFields("total_amount_issued").Result = total_amount_issued
		.FormFields("worker_signature").Result = worker_signature
		.FormFields("Worker_title").Result = Worker_title
		.FormFields("Phone_number_editbox").Result = Phone_number_editbox
		.FormFields("Date_complete").Result = Date_complete

                'Ect to fill in all the blanks in the documents 
End With


script_end_procedure("")









