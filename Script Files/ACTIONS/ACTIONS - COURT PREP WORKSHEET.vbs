'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - COURT PREP WORKSHEET.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog


BeginDialog CAAD_dialog, 0, 0, 176, 85, "Court Prep Dialog"
  EditBox 70, 45, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 165, 25, "Take time now to review your case for court prep.  When you are done reviewing the case, click OK to enter a CAAD note."
  Text 5, 50, 60, 10, "Worker signature:"
EndDialog

'FUNCTIONS----------------------------------------------------------------------------------------------------


'***************************************************************************************************************
' This is a custom function to change the format of a participant name.  The parameter is a string with the
' client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".
FUNCTION change_client_name_to_FML(client_name)
	client_name = trim(client_name)
	length = len(client_name)
	position = InStr(client_name, ", ")
	last_name = Left(client_name, position-1)
	first_name = Right(client_name, length-position-1)
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
END FUNCTION
'***************************************************************************************************************
' This is a custom function to return a listing of a participant's open cases where the participant is a particular role, excluding a particular case.
' The resulting list shows the participants' other cases.  Participant MCI #, role, and case number to exclude are parameters.

FUNCTION participant_case_browse_excluding(mci_number, role, case_number_to_exclude)
	CALL navigate_to_PRISM_screen(main)
	if role = "CP" then
		browse = "CPCB"
	else
		browse = "NCCB"
	end if
	EMWriteScreen browse, 21, 18
	EMWriteScreen mci_number, 20, 6
	transmit


	browse_row = 7
	active_case = false

	DO
		EMReadScreen end_of_data, 11, browse_row, 32

		if end_of_data <> "End of Data" then

			EMReadScreen browse_role, 3, browse_row, 8
			EMReadScreen browse_stat, 3, browse_row, 68
			EMReadScreen browse_case_num_first, 10, browse_row, 15
			EMReadScreen browse_case_num_second, 2, browse_row, 26
			browse_case_number = browse_case_num_first & "-" & browse_case_num_second
			'Check the role and case status - we only want active cases
			If browse_role = role and browse_stat = "OPN" and browse_case_number <> case_number_to_exclude then
				active_case = true
				EMWriteScreen "S", browse_row, 4
				transmit
				EMSetCursor 5, 53
				PF1
				EMReadScreen browse_county, 40, 9, 35
				PF3
				CALL navigate_to_PRISM_screen("cafs")
				transmit
				EMReadScreen mo_accrual, 14, 9, 25
				EMReadScreen mo_nonaccrual, 14, 10, 25
				EMReadScreen total_arrears, 14, 12, 64
				case_info = case_info + "Agency: " & browse_county & chr(13)& "Case Number: " & browse_case_number & chr(13) & "Monthly Accrual: " & FormatCurrency(mo_accrual)& chr(13) & "Monthly Non-accrual: " & FormatCurrency(mo_nonaccrual) & chr(13) & "Total Case Arrears Balance: " & FormatCurrency(total_arrears) & chr(13) & chr(13)
				CALL navigate_to_PRISM_screen(browse)
				EMWriteScreen mci_number, 20, 6
				transmit
			END IF
			browse_row = browse_row + 1
			IF NCID_row = 19 THEN
				PF8
				NCID_row = 8
			END IF
			If active_case = false then
				case_info = "There were no active cases with MCI # " & mci_number & " as a " & role &" found."
			End if
		END IF
	LOOP UNTIL end_of_data = "End of Data"

	participant_case_browse_excluding = case_info 'To make this a return function, this statement must set the value of the function name
END FUNCTION
'**************************************************************************************************************
FUNCTION participant_case_browse(mci_number, role)
	CALL navigate_to_PRISM_screen(main)
	if role = "CP" then
		browse = "CPCB"
	else
		browse = "NCCB"
	end if
	EMWriteScreen browse, 21, 18
	EMWriteScreen mci_number, 20, 6
	transmit


	browse_row = 7
	active_case = false

	DO
		EMReadScreen end_of_data, 11, browse_row, 32

		if end_of_data <> "End of Data" then

			EMReadScreen browse_role, 3, browse_row, 8
			EMReadScreen browse_stat, 3, browse_row, 68
			EMReadScreen browse_case_num_first, 10, browse_row, 15
			EMReadScreen browse_case_num_second, 2, browse_row, 26

			'Check the role and case status - we only want active cases
			If browse_role = role and browse_stat = "OPN" then
				active_case = true
				EMWriteScreen "S", browse_row, 4
				transmit
				EMSetCursor 5, 53
				PF1
				EMReadScreen browse_county, 40, 9, 35
				PF3
				CALL navigate_to_PRISM_screen("cafs")
				transmit
				EMReadScreen mo_accrual, 14, 9, 25
				EMReadScreen mo_nonaccrual, 14, 10, 25
				EMReadScreen total_arrears, 14, 12, 64
				case_info = case_info + "County: " & browse_county & chr(13)& "Case Number: " & browse_case_num_first & "-" & browse_case_num_second & chr(13) & "Monthly Accrual: " & FormatCurrency(mo_accrual)& chr(13) & "Monthly Non-accrual: " & FormatCurrency(mo_nonaccrual) & chr(13) & "Total Case Arrears Balance: " & FormatCurrency(total_arrears) & chr(13) & chr(13)
				CALL navigate_to_PRISM_screen(browse)
				EMWriteScreen mci_number, 20, 6
				transmit
			END IF
			browse_row = browse_row + 1
			IF NCID_row = 19 THEN
				PF8
				NCID_row = 8
			END IF
			If active_case = false then
				case_info = "There were no active cases with MCI # " & mci_number & " as a " & role &" found."
			End if
		END IF
	LOOP UNTIL end_of_data = "End of Data"


	participant_case_browse = case_info 'To make this a return function, this statement must set the value of the function name
END FUNCTION

'***************************************************************************************************************
FUNCTION create_NCID_variable(NCID)
	CALL navigate_to_PRISM_screen("NCID")
	CALL write_value_and_transmit("B", 3, 29)

	NCID = "Employment history: " & chr(13) & chr(13)
	employer_found = false
	NCID_row = 8
	DO
		EMReadScreen end_of_data, 11, NCID_row, 32
		EMReadScreen employer, 30, NCID_row, 51
		employer = trim(employer)

		EMReadScreen begin_date, 8, NCID_row, 7
		EMReadScreen end_date, 8, NCID_row, 16
		end_date = trim(end_date)
		IF end_of_data <> "End of Data" AND end_date = "" THEN
			NCID = NCID & "Currently employed at: " & employer & chr(13) & " Start Date: " & begin_date & " " & chr(13)
			employer_found = true
		ELSE
			IF trim(employer) <> "" then
				NCID = NCID & "Previously employed at: " & employer & chr(13) & " Start Date: " & begin_date & " to End Date: " & end_date & "; "& chr(13)
				employer_found = true
			END IF
			IF employer_found = false then
				NCID = NCID & "Unknown"
			END IF
		END IF
		NCID_row = NCID_row + 1
		IF NCID_row = 19 THEN
			PF8
			NCID_row = 8
		END IF

	LOOP UNTIL end_of_data = "End of Data"
	create_NCID_variable = NCID	'To make this a return function, this statement must set the value of the function name
END FUNCTION
'***************************************************************************************************************
FUNCTION create_PALC_variable(PALC)
	CALL navigate_to_PRISM_screen("PALC")
	EMWriteScreen date, 20, 49
	transmit
	CALL read_PALC_last_payment_date(last_payment_date)
	PALC = PALC & "Last Payment Date: " & last_payment_date & "; "
	CALL read_PALC_payment_type(payment_type)
	PALC = PALC & "Payment Type: " & payment_type & "; "
	CALL read_PALC_payment_amount(payment_amount)
	PALC = PALC & "Payment Amount: " & payment_amount & "; "
	CALL read_PALC_alloc_amount(alloc_amount)
	PALC = PALC & "Case Allocated Amount: " & alloc_amount
	create_PALC_variable = PALC  'To make this a return function, this statement must set the value of the function name
END FUNCTION
'***************************************************************************************************************
FUNCTION read_PALC_payment_type(payment_type)
	CALL navigate_to_PRISM_screen("PALC")
	EMWriteScreen date, 20, 49
	transmit
	EMReadScreen payment_type, 3, 9, 25
END FUNCTION
'***************************************************************************************************************
FUNCTION read_PALC_last_payment_date(last_payment_date)
	CALL navigate_to_PRISM_screen("PALC")
	EMWriteScreen date, 20, 49
	transmit
	EMWriteScreen "D", 9, 5
	transmit
	EMReadScreen last_payment_date, 8, 13, 37
	PF3
END FUNCTION
'***************************************************************************************************************
FUNCTION read_PALC_payment_amount(payment_amount)
	CALL navigate_to_PRISM_screen("PALC")
	EMWriteScreen date, 20, 49
	transmit
	EMReadScreen payment_amount, 13, 9, 29
	payment_amount = trim(payment_amount)
END FUNCTION
'***************************************************************************************************************
FUNCTION read_PALC_alloc_amount(alloc_amount)
	CALL navigate_to_PRISM_screen("PALC")
	EMWriteScreen date, 20, 49
	transmit
	EMReadScreen alloc_amount, 12, 9, 68
	alloc_amount = trim(alloc_amount)
END FUNCTION
'***************************************************************************************************************
FUNCTION create_PAPD_variable(PAPD)

	CALL navigate_to_PRISM_screen("PAPD")
	CALL write_value_and_transmit("B", 3, 29)

	PAPD_row = 8
	DO
		EMReadScreen end_of_data, 11, PAPD_row, 32
		EMReadScreen papd_remedy, 3, PAPD_row, 30
		EMReadScreen pay_plan_begin, 8, PAPD_row, 47
		EMReadScreen pay_plan_end, 8, PAPD_row, 58
			pay_plan_end = replace(pay_plan_end, " ", "")
		IF end_of_data <> "End of Data" AND pay_plan_end = "" THEN
			EMSetCursor PAPD_row, 30
			transmit

			CALL find_variable("TTl Amt: ", ttl_due, 13)
				ttl_due = trim(ttl_due)
			CALL find_variable("Delq Amt: ", delinquent, 13)
				delinquent = trim(delinquent)

			PAPD = PAPD & "Remedy: " & papd_remedy & "; " & "Begin Date: " & pay_plan_begin & "; " & "Total Due: " & ttl_due & "; " & "Delinq. Amount: " & delinquent & ";"

			CALL write_value_and_transmit("B", 3, 29)
		END IF
		PAPD_row = PAPD_row + 1
	LOOP UNTIL end_of_data = "End of Data"
	create_PAPD_variable = PAPD
END FUNCTION


'***************************************************************************************************************


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
call navigate_to_PRISM_screen("CAPS")

'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it

EMReadScreen CP_name, 30, 6, 12
CP_name = trim(CP_name)
CP_name = change_client_name_to_FML(CP_name)

EMReadScreen NCP_name, 30, 7, 12
NCP_name = trim(NCP_name)
NCP_name = change_client_name_to_FML(NCP_name)

EMReadScreen mci_number, 10, 8, 11


'Getting worker info
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
worker_name = trim(worker_name)
worker_name = change_client_name_to_FML(worker_name)

EMReadScreen worker_phone, 12, 8, 35
PF3

'Cleaning up worker info

'EMPLOYMENT INFO

NCID = create_NCID_Variable(NCID)


'SAFETY CONCERNS INFO
call navigate_to_PRISM_screen("GCSC")
EMReadScreen CP_concerns, 1, 12, 24
EMReadScreen NCP_concerns, 1, 13, 24
EMReadScreen GC_code, 12, 16, 24

'FINANCIAL SUMMARY
call navigate_to_PRISM_screen("CAFS")
EMReadScreen total_arrears, 14, 12, 64
EMReadScreen PA_arrears, 14, 11, 64
EMReadScreen NPA_arrears, 14, 10, 64
call navigate_to_PRISM_screen("NCOL")
EMWriteScreen "CCH", 20, 39
transmit
EMReadScreen CCH_amount, 9, 9, 36
EMWriteScreen "CCC", 20, 39
transmit
EMReadScreen CCC_amount, 9, 9, 36
EMWriteScreen "CMI", 20, 39
transmit
EMReadScreen CMI_amount, 9, 9, 36
EMWriteScreen "CMS", 20, 39
transmit
EMReadScreen CMS_amount, 9, 9, 36
EMWriteScreen "JCH", 20, 39
transmit
EMReadScreen JCH_amount, 9, 9, 36



'EMANCIPATION INFO
call navigate_to_PRISM_screen("CHIC")
'Getting all child/DOB info
PRISM_row = 8
Do
	EMReadScreen end_of_data, 11, PRISM_row, 32
 	If end_of_data <> "End of Data" then
		EMReadScreen child_name, 30, PRISM_row, 7	'reading name
		child_name = trim(child_name)		'removing spaces from beginning and end
		child_name = change_client_name_to_FML(child_name)

	 	EMWriteScreen "D", PRISM_row, 4
		transmit
		EMReadScreen child_DOB, 10, 11, 47	'reading DOB
		EMReadScreen child_18th_bday, 10, 19, 17
		CHICS_kids = CHICS_kids & child_name & chr(13)& " DOB: " & child_DOB &  chr(13) & "18th B-day: "& child_18th_bday & chr(13) & chr(13) 		'If there's a name, add to the CHICS_kids variable
		PF3
		PRISM_row = PRISM_row + 1					'increase the PRISM row
		If PRISM_row = 19 then						'If we're on row 19, go to the next page
			PF8
			PRISM_row = 8
		End if
	END IF
Loop until end_of_data = "End of Data"


'PAPD INFO
PAPD = create_PAPD_variable(PAPD)

'PAYMENT INFO
PALC = create_PALC_variable(PALC)

'PA PROGRAMS FOR CP/CHILD

'PA PROGRAMS FOR NCP

'NCP'S OTHER CASES
case_info = participant_case_browse_excluding(mci_number, "NCP", PRISM_case_number)

call navigate_to_PRISM_screen("cast")
'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
'Do
'	Dialog CS_intake_dialog
'	If buttonpressed = 0 then stopscript
'	transmit
'	EMReadScreen PRISM_check, 5, 1, 36
'	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
'Loop until PRISM_check = "PRISM"


'Creating the Word application object and making it visible
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

'Creating document
set objDoc = objWord.Documents.Add(word_documents_folder_path & "Court Prep Worksheet.dotx")
set objSelection = objWord.Selection
	With objDoc
		.FormFields("NCP_NameFML").Result = NCP_name
		.FormFields("CP_NameFML").Result = CP_name
		.FormFields("CSO_NameFML").Result = worker_name
		.FormFields("CSO_Phone").Result = worker_phone
		.FormFields("PRISM_Number").Result = PRISM_case_number
		.FormFields("MCI_Number").Result = mci_number

		.FormFields("Safety_Info").Select
		objSelection.TypeText "CP Concern - " & CP_concerns & vbCr
		objSelection.TypeText "NCP Concern - " & NCP_concerns & vbCr
		objSelection.TypeText "Good Case: " & GC_code


		if inStr(Cstr(CCH_amount), "Data") <= 0 then
			.FormFields("CCH").Result = CCH_amount
		else
			.FormFields("CCH").Select
			objSelection.TypeText "No active CCH obligation."
		end if
		if instr(Cstr(CCC_amount), "Data") <= 0 then
			.FormFields("CCC").Result = CCC_amount
		else
			.FormFields("CCC").Select
			objSelection.TypeText "No active CCC obligation."
		end if
		if inStr(Cstr(CMI_amount), "Data") <= 0 then
			.FormFields("CMICMS").Result = "CMI: " & CMI_amount
		elseif instr(Cstr(CMS_amount), "Data") <= 0 then
			.FormFields("CMICMS").Result = "CMS: " & CMS_amount
		else
			.FormFields("CMICMS").Select
			objSelection.TypeText "No active CMI or CMS obligation."
		end if
		if instr(Cstr(JCH_amount), "Data") <= 0 then
			.FormFields("JCH").Result = JCH_amount
		else
			.FormFields("JCH").Select
			objSelection.TypeText "No active JCH obligation."
		end if
		.FormFields("arrears").Result = total_arrears
		.FormFields("pa_arrears").Result = PA_arrears
		.FormFields("npa_arrears").Result = NPA_arrears

		.FormFields("Emancipation_Info").Select
		objSelection.TypeText CHICS_kids

		.FormFields("PAPD_Info").Select
	'	objSelection.TypeText PAPD

		.FormFields("Payment_Info").Select
		objSelection.TypeText PALC

		.FormFields("Employment_Info").Select
		objSelection.TypeText NCID

		.FormFields("Other_cases").Select
		objSelection.TypeText case_info

	End With
'Shows case note dialog
Do
	Do
		Dialog CAAD_dialog
		If buttonpressed = 0 then stopscript

		If worker_signature = ""  then MsgBox "You must enter your signature!"
	Loop until worker_signature <> ""

	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


	'Going to CAAD, adding a new note
	call navigate_to_PRISM_screen("CAAD")
	PF5
	EMWriteScreen "A", 8, 5
	EMReadScreen case_activity_detail, 20, 2, 29

	If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")


	'Setting the type
	EMWriteScreen "FREE", 4, 54

	'Setting cursor in write area and writing note details
	EMSetCursor 16, 4
	call write_variable_in_CAAD("* Used court prep worksheet script and reviewed case for court prep.")

	call write_variable_in_CAAD("---")
	call write_variable_in_CAAD(worker_signature)


script_end_procedure("")
