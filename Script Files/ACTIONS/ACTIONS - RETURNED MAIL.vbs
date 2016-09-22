'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - RETURNED MAIL.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 60
STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------

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


'Calling the Returned Mail Dialog--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog returned_mail_dialog, 0, 0, 441, 195, "Returned Mail Received"
  EditBox 85, 5, 95, 15, prism_number
  CheckBox 100, 30, 25, 10, "CP", rm_cp_checkbox
  CheckBox 135, 30, 50, 10, "NCP/ALF", rm_ncp_checkbox
  EditBox 225, 25, 80, 15, rm_other
  DropListBox 70, 50, 150, 15, "Select one..."+chr(9)+"Update to Unknown"+chr(9)+"Update to New Forwarding Address", updated_ADDR
  EditBox 230, 60, 75, 15, date_received
  EditBox 75, 80, 125, 15, new_ADDR
  EditBox 75, 100, 120, 15, new_CITY
  EditBox 75, 120, 25, 15, new_STATE
  EditBox 160, 120, 45, 15, new_ZIP
  DropListBox 260, 80, 115, 15, "Select one..."+chr(9)+"APP - Application for CS Services"+chr(9)+"COO - Court Order"+chr(9)+"COU - Court"+chr(9)+"CRB - Credit Bureau"+chr(9)+"CUP - Custodial Parent"+chr(9)+"DES - Dept Economic Security"+chr(9)+"DIR - City Directory"+chr(9)+"DOC - Dept of Corrections"+chr(9)+"DPS - Dept Public Safety"+chr(9)+"EMP - Employer"+chr(9)+"INT - Interstate"+chr(9)+"MAX - Maxis"+chr(9)+"NCP - Non Custodial Parent"+chr(9)+"OTH - Other"+chr(9)+"POS - US Postal Service", source_code
  ComboBox 290, 105, 140, 15, "Select One or Leave Blank..."+chr(9)+"MDA - Mail delivered as addressed"+chr(9)+"MFE - Moved, Forwarding Expired"+chr(9)+"MNF - Moved, No Forwarding Address"+chr(9)+"NKA - Not known as Addressed"+chr(9)+"NSA - No such Address"+chr(9)+"OTH - Other"+chr(9)+"PGA - Post Office Gave New Address", postal_resp_code
  EditBox 100, 150, 240, 15, misc_notes
  EditBox 75, 170, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 235, 170, 50, 15
    CancelButton 290, 170, 50, 15
  Text 200, 30, 20, 10, "Other:"
  Text 10, 50, 60, 10, "Update Address?"
  Text 230, 50, 80, 10, "Effective/Verified Date:"
  GroupBox 5, 65, 215, 80, "New Address Info (If given by Post Office)"
  Text 20, 85, 50, 10, "Street Address:"
  Text 55, 105, 20, 10, "City:"
  Text 50, 125, 20, 10, "State:"
  Text 125, 125, 35, 10, "Zip Code:"
  Text 230, 85, 30, 10, "Source:"
  Text 230, 110, 60, 10, "Postal Response:"
  Text 10, 155, 90, 10, "Misc notes/Actions Taken:"
  Text 10, 175, 65, 10, "Worker Signature:"
  Text 10, 10, 70, 10, "PRISM Case Number:"
  Text 10, 30, 85, 10, "Returned Mail Rec'd for:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to Bluezone
EMConnect ""			

'Checks for prism case number and navigates to CAST for that case
CALL check_for_prism(TRUE)
CALL PRISM_case_number_finder(prism_number)
CALL navigate_to_PRISM_screen("CAST")
EMReadScreen function_code, 2, 5, 78

'The script will not run unless the mandatory fields are completed
DO
	err_msg = ""
	Dialog returned_mail_dialog
	IF ButtonPressed = 0 THEN StopScript		                                       
	CALL PRISM_case_number_validation(prism_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF forwarding_addr = "Select One..." THEN err_msg = err_msg & vbNewline & "You must answer if there was a forwarding address given!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"       
	IF date_received = "" THEN err_msg = err_msg & vbNewline & "You must enter a effective/verified date."            
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

'Cutting postal response to three characters
source= left(source_code, 3)
postal_resp=left(postal_resp_code, 3)


'Navigates to CPDD, NCDD or CAAD note for other address
IF rm_cp_checkbox = CHECKED THEN 
'Do we need to add a new address or set address to unknown?
	IF updated_ADDR = "Update to New Forwarding Address" THEN
		CALL navigate_to_PRISM_screen("CPDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit
'Erasing the current address in PRISM
		EMWritescreen "M", 3, 29	
		EMWritescreen date_received, 10, 18
		EMWritescreen "N", 10, 46	
		EMSetCursor 14, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 15, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 16, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 39	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 50	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 56	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 69	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 7	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 38	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 62	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
'Adding the new forwarding address in PRISM		
		CALL navigate_to_PRISM_screen("CPDD")
		EMwritescreen "M", 3, 29
		EMWritescreen date_received, 10, 18
		EMWritescreen "Y", 10, 46
		EMwritescreen new_addr, 15, 11
		EMWritescreen new_CITY, 17, 11
		EMWritescreen new_STATE, 17, 39
		EMWritescreen new_ZIP, 17, 50
		EMWritescreen date_received, 19, 7
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

'Shows error message if one exists
		EMReadScreen standardization_msg, 6, 4, 35
		IF standardization_msg = "Code-1" THEN
			EMReadscreen error_msg, 29, 12, 25
			IF error_msg <> "Address has been standardized" THEN
				PF6
				PF3	
				Msgbox "PRISM reports this message: " & trim(error_msg) & ". Please review and/or update the address if applicable! The R0011 CAAD note will not be entered. Script will now stop."
				stopscript
			END IF	
			PF6
		END IF

		IF function_code = "OL" or function_code = "ON" or function_code = "PL" or function_code = "PN" THEN
			MsgBox "** Review case to see if maintaining county request needs to be made **"
		END IF
'Navigates to CAAD and enters the CAAD note for returned mail
CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen date_received, 4, 37
	EMWritescreen "R0011", 4, 54
	EMSetCursor 16, 4
	CALL write_variable_in_CAAD(misc_notes)
	CALL write_variable_in_CAAD("---" & worker_signature)
	transmit	

'Erases the address
	ELSEIF updated_ADDR = "Update to Unknown" THEN
		CALL navigate_to_PRISM_screen("CPDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMWritescreen "M", 3, 29	
		EMWritescreen date_received, 10, 18
		EMWritescreen "N", 10, 46	
		EMSetCursor 14, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 15, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 16, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 39	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 50	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 56	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 69	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 7	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 38	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 62	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		transmit
	END IF
END IF

'Navigates to CAAD to write a case note
IF rm_ncp_checkbox = CHECKED THEN
	CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen date_received, 4, 37
	EMWritescreen "R0010", 4, 54
	EMSetCursor 16, 4
	CALL write_variable_in_CAAD(misc_notes)
	CALL write_variable_in_CAAD("---" & worker_signature)
	transmit
'Erases the old address
	IF updated_ADDR = "Update to New Forwarding Address" THEN
		CALL navigate_to_PRISM_screen("NCDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMWritescreen "M", 3, 29	
		EMWritescreen date_received, 10, 18
		EMWritescreen "N", 10, 46	
		EMSetCursor 14, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 15, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 16, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 39	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 50	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 56	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 69	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 7	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 38	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 62	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		
'Enteres the new forwarding address
		EMwritescreen "M", 3, 29
		EMWritescreen date_received, 10, 18
		EMWritescreen "Y", 10, 46
		EMwritescreen new_addr, 15, 11
		EMWritescreen new_CITY, 17, 11
		EMWritescreen new_STATE, 17, 39
		EMWritescreen new_ZIP, 17, 50
		EMWritescreen date_received, 19, 7
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit
'Shows error message if there is one		
		EMReadScreen standardization_msg, 6, 4, 35
		IF standardization_msg = "Code-1" THEN
			EMReadscreen error_msg, 29, 12, 25
			IF trim(error_msg) <> "Address has been standardized" THEN
				PF6
				PF3	
				Msgbox "PRISM reports this message: " & trim(error_msg) & ". Please review and/or update the address if applicable! The R0010 CAAD note will not be entered. Script will now stop"
				stopscript
			END IF	
			PF6
		END IF
'Navigates to caad and enters the returned mail note		
CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen date_received, 4, 37
	EMWritescreen "R0010", 4, 54
	EMSetCursor 16, 4
	CALL write_variable_in_CAAD(misc_notes)
	CALL write_variable_in_CAAD("---" & worker_signature)
	transmit		

'Erases address and saves it
	ELSEIF updated_ADDR = "Update to Unknown" THEN
		CALL navigate_to_PRISM_screen("NCDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMWritescreen "M", 3, 29	
		EMWritescreen date_received, 10, 18
		EMWritescreen "N", 10, 46	
		EMSetCursor 14, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 15, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 16, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 39	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 50	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 56	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 69	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 7	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 38	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 62	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		transmit
	END IF
END IF

'Enters the caad note for other returned mail
IF rm_other <> "" THEN
	CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen "R0012", 4, 54
	EMSetCursor 16, 4
	CALL write_variable_in_CAAD(rm_other)
	CALL write_variable_in_CAAD(misc_notes)
	CALL write_variable_in_CAAD("---" & worker_signature)
	transmit
END IF


script_end_procedure("")
