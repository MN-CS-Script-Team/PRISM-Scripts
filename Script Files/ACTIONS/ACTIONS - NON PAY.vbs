'GATHERING STATS---------------------------------------------------------------------
'name_of_script = "ACTIONS - NONPAY LTR.vbs"
'start_time = timer
'
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



'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 85, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog
'NONPAY LTR DIAL0G -
BeginDialog NONPAY_LTR_DIALOG, 0, 0, 176, 146, "NONPAY LTR Dialog"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 60, 10, "Non Pay Ltr", NonPay_button
    PushButton 10, 60, 100, 10, "Non Compliance w Pay Plan", PAPD_button
    PushButton 10, 110, 100, 10, "Send CP Initial Docs", Initial_docs_button
    CancelButton 0, 0, 0, 0
  Text 10, 40, 140, 20, "Send DL Non Compliance Ltr"
  Text 10, 0, 140, 20, "Send Nonpay Letter and E9685 CAAD"
  Text 10, 80, 150, 25, "Send CP Initial Contempt Ltr, Role of CAO, Special Services Assessment and M2123 CAAD"
  ButtonGroup ButtonPressed
    PushButton 10, 130, 70, 10, "Cancel", Cancel_button
EndDialog



'************FUNCTIONS **************************
FUNCTION fix_read_data (search_string) 
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
END FUNCTION

FUNCTION send_non_compliance_dord
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "A", 3, 29
	EMWriteScreen "        ", 4, 50
	EMWriteScreen "       ", 4, 59
	EMWriteScreen "F0919", 6, 36
	transmit
END FUNCTION


FUNCTION send_non_pay_memo
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "A", 3, 29
	'-----Selecting the form
	EMWriteScreen "F0104", 6, 36
	'-----Selecting the recipient
	EMWriteScreen "NCP", 11, 51
	EMWriteScreen "        ", 4, 50
	EMWriteScreen "       ", 4, 59
	transmit

	EMSendKey "<PF14>"
	EMWaitReady 0, 0

	EMWriteScreen "U", 20, 14
	transmit

	dord_row = 7
	DO
		EMWriteScreen "S", dord_row, 5
		dord_row = dord_row + 1
	LOOP UNTIL dord_row = 19
	transmit
	
	EMWriteScreen "As you are aware, you have a court ordered obligation to pay ", 16, 15
	transmit
	EMWriteScreen "child support. Last payment of: $" & last_payment_amount & " was received", 16, 15
	transmit
	EMWriteScreen "on: " & last_payment_date & ". All court ordered obligations must", 16, 15
	transmit
	EMWriteScreen "be paid during the month in which they are due. Failure to", 16, 15
	transmit
	EMWriteScreen "pay your court ordered support obligation can result in", 16, 15
	transmit
	EMWriteScreen "actions such as: suspension of driver's license, seizure", 16, 15
	transmit
	EMWriteScreen "of funds held in a financial institution, denial of", 16, 15
	transmit
	EMWriteScreen "passport, suspension of recreational licenses such as", 16, 15
	transmit
	EMWriteScreen "fishing and hunting, interception of tax refunds, suspension of any", 16, 15
	transmit
	EMWriteScreen "professional license you may hold, reporting of your arrears", 16, 15
	transmit
	EMWriteScreen "balance to the major credit reporting agencies and possible ", 16, 15
	transmit
	EMWriteScreen "court action for non-payment of support. Please contact me ", 16, 15
	transmit
	transmit

	dord_row = 7
	DO
		EMWriteScreen "S", dord_row, 5
		dord_row = dord_row + 1
	LOOP UNTIL dord_row = 13
	transmit

	EMWriteScreen "to discuss your employment status or sources of income.", 16, 15
	transmit
	EMWriteScreen "Please make a payment today. Your current arrears balance", 16, 15
	transmit
	EMWriteScreen "is $" & arrears_balance & ". If you have any questions or concerns", 16, 15
	transmit
	EMWriteScreen "regarding your support obligation, please contact me at the", 16, 15
	transmit
	EMWriteScreen "number listed below. Here's the link to making payments", 16, 15
	transmit
	EMWriteScreen "online:  http://www.childsupport.dhs.state.mn.us", 16, 15
	transmit
	
	PF3

	EMWriteScreen "M", 3, 29
	transmit	

	PF9
	PF3	
END FUNCTION

FUNCTION add_caad_code(CAAD_code)
	CALL navigate_to_PRISM_screen("CAAD")	
	PF5
	EMWriteScreen CAAD_code, 4, 54
END FUNCTION

'***************************

'Connecting to BlueZone
EMConnect ""
CALL check_for_PRISM(True)
call PRISM_case_number_finder(PRISM_case_number)

CALL navigate_to_PRISM_screen("PAPL")
EMReadScreen last_payment_date, 8, 7, 40   
EMReadScreen last_payment_amount, 7, 7, 53
CALL navigate_to_PRISM_screen("CAFS")
EMReadScreen arrears_balance, 9, 12, 69
'Case number display dialog
Do
	
	Dialog case_number_dialog
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
Loop until case_number_valid = True
			
	Do
		EMReadScreen PRISM_check, 5, 1, 36
		If PRISM_check <> "PRISM" then MsgBox "You appear to have timed out, or are out of PRISM. Navigate to PRISM and try again."
	Loop until PRISM_check = "PRISM"
	
	Dialog NONPAY_LTR_DIALOG
	
	
	IF ButtonPressed = Cancel_button THEN stopscript
	If ButtonPressed = NonPay_button then 
		CALL send_non_pay_memo
		purge_msg = MsgBox ("Do you want to purge this worklist item?", vbYesNo)
		IF purge_msg = vbYes THEN 
			DO	
				CALL navigate_to_PRISM_screen("CAWT")			
				EMWriteScreen PRISM_case_number, 20, 8
				EMWriteScreen "E0002", 20, 29
				transmit
				EMReadScreen cawt_type, 5, 8, 8
				If cawt_type = "E0002" then 
					EMWriteScreen "P", 8, 4
					transmit
					transmit
					PF3
				End If	
			Loop until cawt_type <> "E0002"
		END IF
		CALL add_caad_code("E9685")
		script_end_procedure("The script has sent the requested DORD document and is now waiting for you to transmit to confirm the CAAD Note.")
	End If
	If ButtonPressed = PAPD_button then 
		CALL send_non_compliance_dord
		purge_msg = MsgBox ("Do you want to purge this worklist item?", vbYesNo)
		IF purge_msg = vbYes THEN 
			Do
				CALL navigate_to_PRISM_screen("CAWT")
				EMWriteScreen PRISM_case_number, 20, 8
				EMWriteScreen "E4111", 20, 29
				transmit
				EMReadScreen cawt_type, 5, 8, 8
				
				If cawt_type = "E4111" then
					EMWriteScreen "P", 8, 4
					transmit
					transmit
					PF3
				End If
			Loop until cawt_type <> "E4111"
		END if	
		script_end_procedure("The script has sent the requested DORD document.")
	END IF
	If ButtonPressed = Initial_docs_button then

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

		'Get information to pull into documents
		EMReadScreen NCP_MCI, 10, 8, 11 
		EMReadScreen CP_MCI, 10, 4, 8 	
		
		'Getting worker info for case note
		EMSetCursor 5, 53
		PF1
		EMReadScreen worker_name, 27, 6, 50
		EMReadScreen worker_phone, 12, 8, 35
		PF3

		'Cleaning up worker info
		worker_name = trim(worker_name)
		call fix_case(worker_name, 1)
		worker_name = change_client_name_to_FML(worker_name)		


		'NCP Name
		call navigate_to_PRISM_screen("NCDE")
		EMWriteScreen NCP_MCI, 4, 7
		EMReadScreen NCP_F, 12, 8, 34
		EMReadScreen NCP_M, 12, 8, 56
		EMReadScreen NCP_L, 17, 8, 8
		EMReadScreen NCP_S, 3, 8, 74
		NCP_name = fix_read_data(NCP_F) & " " & fix_read_data(NCP_M) & " " & fix_read_data(NCP_L)	
		If trim(NCP_S) <> "" then NCP_name = NCP_Name & " " & ucase(fix_read_data(NCP_S))
		NCP_name = trim(NCP_name)

		'CP Name											
		call navigate_to_PRISM_screen("CPDE")
		EMWriteScreen CP_MCI, 4, 7
		EMReadScreen CP_F, 12, 8, 34
		EMReadScreen CP_M, 12, 8, 56
		EMReadScreen CP_L, 17, 8, 8
		EMReadScreen CP_S, 3, 8, 74
		CP_name = fix_read_data(CP_F) & " " & fix_read_data(CP_M) & " " & fix_read_data(CP_L)	
		If trim(CP_S) <> "" then CP_name = CP_Name & " " & ucase(fix_read_data(CP_S))
		CP_name = trim(CP_Name)

		'CP Address
		'Navigating to CPDD to pull address info
		call navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_street_line_1, 30, 15, 11
		EMReadScreen cp_street_line_2, 30, 16, 11
		EMReadScreen cp_city_state_zip, 49, 17, 11

		'Cleaning up address info
		cp_street_line_1 = replace(cp_street_line_1, "_", "")
		call fix_case(cp_street_line_1, 1)
		cp_street_line_2 = replace(cp_street_line_2, "_", "")
		if trim (cp_street_line_2) <> "" then
			cp_address = cp_street_line_1 & chr(13) & cp_street_line_2
		else
			cp_address = cp_street_line_1
		end if
		call fix_case(cp_street_line_2, 1)
		cp_city_state_zip = replace(replace(replace(cp_city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
		call fix_case(cp_city_state_zip, 2)

		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
		set objDoc = objWord.Documents.Add(word_documents_folder_path & "Initial Contempt Docs.dotm")
		With objDoc
		.FormFields("CP").Result = CP_name
		.FormFields("CP1").Result = CP_name
		.FormFields("CP_Address").Result = CP_address
		.FormFields("CP_CSZ").Result = cp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CaseNumber1").Result = PRISM_case_number
		.FormFields("NCP").Result = NCP_name
		.FormFields("NCP1").Result = NCP_name
		.FormFields("NCP2").Result = NCP_name
		.FormFields("NCP3").Result = NCP_name
		.FormFields("party_name").Result = CP_name
		.FormFields("party_name1").Result = CP_name
		.FormFields("party_name2").Result = CP_name
		.FormFields("party_name3").Result = CP_name
		.FormFields("Address").Result = CP_address
		.FormFields("CSZ").Result = cp_city_state_zip
		.FormFields("Case_Number").Result = PRISM_case_number
		.FormFields("Case_Number1").Result = PRISM_case_number
		.FormFields("other_party").Result = NCP_name
		.FormFields("other_party1").Result = NCP_name
		.FormFields("other_party2").Result = NCP_name
		.FormFields("other_party3").Result = NCP_name
		.FormFields("other_party4").Result = NCP_name
		.FormFields("other_party5").Result = NCP_name
		.FormFields("other_party6").Result = NCP_name
		.FormFields("due_date").Result = dateadd("d", date, 10)
		.FormFields("due_date1").Result = dateadd("d", date, 10)
		.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerPhone").Result = worker_phone
			
	End With
		CALL add_caad_code("M2123")
		EMSetCursor 16, 4
		CALL write_new_line_in_PRISM_case_note("    * CP Initial Contempt Letter, Role of County Attorney, and Special Services Coverletter and Assessment sent to CP.")
		transmit
 end if
