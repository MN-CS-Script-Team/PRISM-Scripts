'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "refer-to-mod.vbs"
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
call changelog_update("03/10/2017", "Initial version.", "Wendy LeVesseur, Anoka County")


'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'------Start of Class definitions--------------------------------------------------------------------------------
'>>>>> CLASSES!!!!!!!!!!!!!!!!!!!!! <<<<<
' This CLASS contains properties used to populate documents
' These properties should not be used for other applications in scripts.
' Every time you call the property, the script will use the class definition to efficiently obtain the requested information.
CLASS doc_info
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' >>>>> CP INFORMATION <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' CP first name
	PUBLIC PROPERTY GET cp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_first_name, 12, 8, 34
		cp_first_name = fix_read_data(cp_first_name)
	END PROPERTY

	' CP last name
	PUBLIC PROPERTY GET cp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_last_name, 17, 8, 8
		cp_last_name = fix_read_data(cp_last_name)
	END PROPERTY	
	
	' CP middle name
	PUBLIC PROPERTY GET cp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_middle_name, 12, 8, 56
		cp_middle_name = fix_read_data(cp_middle_name)
	END PROPERTY
	
	' CP middle initial
	PUBLIC PROPERTY GET cp_middle_initial
		cp_middle_initial = left(cp_middle_name, 1)
	END PROPERTY
	
	' CP suffix
	PUBLIC PROPERTY GET cp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_suffix, 3, 8, 74
		cp_suffix = trim(replace(cp_suffix, "_", ""))
	END PROPERTY
	
	' CP MCI
	PUBLIC PROPERTY GET cp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_mci, 10, 5, 7
	END PROPERTY	
	
	' CP address
	PUBLIC PROPERTY GET cp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadscreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_addr1, 30, 15, 11
			EMReadScreen cp_addr2, 30, 16, 11
			cp_addr2 = replace(cp_addr2, "_", "")
				IF cp_addr2 <> "" THEN
					cp_addr = replace(cp_addr1, "_", "") & ", " & replace(cp_addr2, "_", "")
				ELSE
					cp_addr = replace (cp_addr1, "_", "")
				END IF
			cp_addr = fix_read_data(cp_addr)
		ELSE
			cp_addr = "Unknown Address"
		END IF
	END PROPERTY

	' CP address city
	PUBLIC PROPERTY GET cp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadscreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_city, 20, 17, 11
			cp_city = fix_read_data(cp_city)
		ELSE
		cp_city = "City"
		END IF
	END PROPERTY

	' CP address state
	PUBLIC PROPERTY GET cp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadscreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_state, 2, 17, 39
		ELSE
			cp_state = "State"
		END IF
	END PROPERTY
	
    ' CP address zip code
	PUBLIC PROPERTY GET cp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_zip, 10, 17, 50
			cp_zip = fix_read_data(cp_zip)
		ELSE
		cp_zip = "ZIP"
		END IF
	END PROPERTY
	

	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>>>> NCP Information <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' NCP Name
	PUBLIC PROPERTY GET ncp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_name, 50, 5, 25
		ncp_name = fix_read_data(ncp_name)
	END PROPERTY
	
	' NCP first name
	PUBLIC PROPERTY GET ncp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_first_name, 12, 8, 34
		ncp_first_name = fix_read_data(ncp_first_name)
	END PROPERTY

	' NCP last name
	PUBLIC PROPERTY GET ncp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_last_name, 17, 8, 8
		ncp_last_name = fix_read_data(ncp_last_name)
	END PROPERTY	
	
	' NCP middle name
	PUBLIC PROPERTY GET ncp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_middle_name, 12, 8, 56
		ncp_middle_name = fix_read_data(ncp_middle_name)
	END PROPERTY
	
	' NCP middle initial
	PUBLIC PROPERTY GET ncp_middle_initial
		ncp_middle_initial = left(ncp_middle_name, 1)
	END PROPERTY
	
	' NCP suffix
	PUBLIC PROPERTY GET ncp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_suffix, 3, 8, 74
		ncp_suffix = trim(replace(ncp_suffix, "_", ""))
	END PROPERTY	
	
	' NCP SSN
	PUBLIC PROPERTY GET ncp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_ssn, 11, 6, 7
	END PROPERTY

	' NCP MCI
	PUBLIC PROPERTY GET ncp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_mci, 10, 5, 7
	END PROPERTY	

	' NCP street address
	PUBLIC PROPERTY GET ncp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_addr1, 30, 15, 11
			EMReadScreen ncp_addr2, 30, 16, 11
			ncp_addr2 = replace(ncp_addr2, "_", "")
				IF ncp_addr2 <> "" THEN
					ncp_addr = replace(ncp_addr1, "_", "") & ", " & replace(ncp_addr2, "_", "")
				ELSE
					ncp_addr = replace (ncp_addr1, "_", "")
				END IF
			ncp_addr = fix_read_data(ncp_addr)
		ELSE
			ncp_addr = "Unknown Address"
		END IF
	END PROPERTY

	' NCP address city
	PUBLIC PROPERTY GET ncp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_city, 20, 17, 11
			ncp_city = fix_read_data(ncp_city)
		ELSE
			ncp_city = "City"
		END IF
	END PROPERTY

	' NCP address state
	PUBLIC PROPERTY GET ncp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_state, 2, 17, 39
		ELSE
			ncp_state = "State"
		END IF
	END PROPERTY
    
	' NCP address zip code
	PUBLIC PROPERTY GET ncp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_zip, 10, 17, 50
			ncp_zip = fix_read_data(ncp_zip)
		ELSE
			ncp_zip = "ZIP"
		END IF
	END PROPERTY

' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>> General Information <<<
	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<

	' Case worker name FML
	PUBLIC PROPERTY GET worker_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMSetCursor 5, 56
		PF1
		EMReadScreen worker_name, 30, 6, 50
		worker_name = change_client_name_to_FML(worker_name)
		worker_name = fix_read_data(worker_name)
		transmit
	END PROPERTY
	
	' Case worker phone ###-###-####
	PUBLIC PROPERTY GET worker_phone
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMSetCursor 5, 56
		PF1
		EMReadScreen worker_phone, 12, 8, 35
		transmit
	END PROPERTY	
END CLASS
'------End of Class definitions--------------------------------------------------------------------------------


'DIALOGS----------------------------------------------------------------------------------------------------
'Prism case number selection dialog, pre-populated if script is run from a case-based screen
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

'Dialog for selection actions for the script to take
BeginDialog Mod_info_dialog, 0, 0, 321, 195, "Mod Actions Dialog"
  CheckBox 10, 25, 135, 10, "Employment Verification - NCP", NCP_EVR_check
  CheckBox 10, 40, 135, 10, "Special Services Assessment - NCP", NCP_Special_assessment_check
  CheckBox 180, 25, 135, 10, "Employment Verification - CP", CP_EVR_check
  CheckBox 180, 40, 135, 10, "Special Services Assessment - CP", CP_Special_assessment_check
  CheckBox 180, 55, 110, 10, "Child Care Verification Form", child_care_verification_check
  CheckBox 10, 80, 105, 10, "Send FPLS request for NCP", NCP_NCMR_check
  CheckBox 180, 80, 125, 10, "Send FPLS request for child(ren)", CH_CHMR_check
  CheckBox 10, 95, 165, 10, "Update Notice of Review with Mod Worker Info", update_mod_worker
  CheckBox 180, 95, 125, 10, "Create followup worklist", CAWD_check
  DropListBox 90, 125, 90, 10, "Select one"+chr(9)+"Theresa Hogan"+chr(9)+"Terri Spence-Garski"+chr(9)+"Kelly Rein"+chr(9)+"Karla Wangrud", mod_worker
  EditBox 100, 150, 90, 15, caad_note_text
  ButtonGroup ButtonPressed
    OkButton 205, 170, 50, 15
    CancelButton 260, 170, 50, 15
  GroupBox 0, 10, 310, 60, "Documents to send out:"
  GroupBox 0, 70, 310, 40, "Actions to take:"
  Text 20, 125, 65, 10, "Mod worker name:"
  Text 15, 155, 80, 10, "Optional CAAD note text:"
EndDialog


'Dialog for selecting a legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
   	OkButton 60, 75, 50, 15
    	CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
 
'Connects to BlueZone
EMConnect ""

continue = msgbox ("Run this script after you have entered the REAM for initiating an agency review.  This script can help send out documents " &_
	"in preparation for modification of court orders and send manual locate requests to obtain social security information. Please click OK to continue or Cancel to stop the script.", 1)
IF continue = vbCancel THEN
	script_end_procedure("The script is now ending.  Action has been cancelled.")
END IF	

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then script_end_procedure("The script has stopped.")
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

'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
Do
	err_msg = ""
	Dialog Mod_info_dialog
	If buttonpressed = 0 then script_end_procedure("The script has stopped.")
	IF update_mod_worker = checked THEN
		IF mod_worker = "Select one" THEN
			err_msg = err_msg & vbCr & "* You must select a modification worker."
		END IF
	END IF
	IF err_msg <> "" THEN Msgbox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

CALL check_for_PRISM(False)

'The command below is necessary in order to utilize the doc_info class to efficiently obtain case data.  We are creating an object called "info" that is a member of the "doc_info" class.
set info = new doc_info

'Creating the Word application object (if any of the Word options are selected), and making it visible 
If _
	child_care_verification_check = checked or _
	CP_Special_assessment_check = checked or _
	NCP_Special_assessment_check = checked then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End if

'These are values that the script uses elsewhere.  We can use the class to populate the property values.
CP_name = info.cp_first_name & " " & info.cp_last_name & " " & info.cp_suffix
NCP_name = info.ncp_first_name & " " & info.ncp_last_name & " " & info.ncp_suffix
CP_name = trim(CP_name)
NCP_name = trim(NCP_name)
CP_mci = info.CP_mci
NCP_mci = info.NCP_mci
NCP_ln = info.ncp_last_name
NCP_fn = info.ncp_first_name
cp_address = info.cp_addr
cp_csz =  info.cp_city & ", " & info.cp_state & "  " & info.cp_zip
ncp_address = info.ncp_addr
ncp_csz =  info.ncp_city & ", " & info.ncp_state & "  " & info.ncp_zip
workers_name = info.worker_name
worker_telephone = info.worker_phone
PRISM_left = Left(PRISM_case_number, 10)
PRISM_right = Right(PRISM_case_number, 2)

'These are variables to indicate if documents were sent.
NCP_EVR_SENT = True
CP_EVR_SENT = True
CHMR_SENT = True
NCP_NCMR_SENT = True

'Creating selected Word documents and populating fields with data
If child_care_verification_check = checked then
	set objDoc = objWord.Documents.Add(word_documents_folder_path & "Childcare-Verification-Letter.dotm")
	With objDoc
		.FormFields("CPName").Result = CP_name 
		.FormFields("CPName1").Result = CP_name
		.FormFields("CPName2").Result = CP_name
		.FormFields("CPAddress").Result = cp_address
		.FormFields("CPCSZ").Result = cp_csz
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CaseNumber1").Result = PRISM_case_number
		.FormFields("ReturnDate").Result = dateadd("d", date, 10)
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_telephone
	End With
End if

If CP_special_assessment_check = checked then
	set objDoc = objWord.Documents.Add(word_documents_folder_path & "Special-Services-Assessment-cover-letter-and-assessment.dotm")
	With objDoc
		.FormFields("party_name").Result = CP_name
		.FormFields("party_name1").Result = CP_name
		.FormFields("party_name2").Result = CP_name
		.FormFields("party_name3").Result = CP_name
		.FormFields("Address").Result = cp_address
		.FormFields("CSZ").Result = cp_csz
		.FormFields("Case_Number").Result = PRISM_case_number
		.FormFields("Case_Number1").Result = PRISM_case_number
		.FormFields("other_party").Result = NCP_name
		.FormFields("other_party1").Result = NCP_name
		.FormFields("other_party2").Result = NCP_name
		.FormFields("other_party3").Result = NCP_name
		.FormFields("other_party4").Result = NCP_name
		.FormFields("other_party5").Result = NCP_name
		.FormFields("other_party6").Result = NCP_name
		.FormFields("due_date").Result = dateadd("d", date, 7)
		.FormFields("due_date1").Result = dateadd("d", date, 7)
		'.FormFields("CaseWorkerName").Result = workers_name
		.FormFields("CaseWorkerPhone").Result = worker_telephone
	End With
End if

If NCP_special_assessment_check = checked then
	set objDoc = objWord.Documents.Add(word_documents_folder_path & "Special-Services-Assessment-cover-letter-and-assessment.dotm")
	With objDoc
		.FormFields("party_name").Result = NCP_name
		.FormFields("party_name1").Result = NCP_name
		.FormFields("party_name2").Result = NCP_name
		.FormFields("party_name3").Result = NCP_name
		.FormFields("Address").Result = ncp_address
		.FormFields("CSZ").Result = ncp_csz
		.FormFields("Case_Number").Result = PRISM_case_number
		.FormFields("Case_Number1").Result = PRISM_case_number
		.FormFields("other_party").Result = CP_name
		.FormFields("other_party1").Result = CP_name
		.FormFields("other_party2").Result = CP_name
		.FormFields("other_party3").Result = CP_name
		.FormFields("other_party4").Result = CP_name
		.FormFields("other_party5").Result = CP_name
		.FormFields("other_party6").Result = CP_name
		.FormFields("due_date").Result = dateadd("d", date, 7)
		.FormFields("due_date1").Result = dateadd("d", date, 7)
		'.FormFields("CaseWorkerName").Result = workers_name
		.FormFields("CaseWorkerPhone").Result = worker_telephone
	End With
End if

'Send EVRs for all CP's active employers.

IF CP_EVR_check = checked THEN
	call navigate_to_PRISM_screen("CPID")
	EMWriteScreen "B", 3, 29
	EMWriteScreen CP_MCI, 8, 15
	transmit
	
	pages = 0
	row = 8
	placeholder_string = ""
	DO
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO

		EMReadScreen employer_code, 10, row, 30
		EMReadScreen emp_end_date, 8, row, 16
		IF trim(emp_end_date) = "" THEN
			IF Instr(placeholder_string, employer_code) = 0 THEN
				placeholder_string = placeholder_string & "~~~" & employer_code
				call navigate_to_PRISM_screen("DORD")

				EMWriteScreen "C", 3, 29
				transmit
				EMWriteScreen "A", 3, 29
				EMWriteScreen CP_MCI, 4, 15
				EMWriteScreen "  ", 4, 26
				EMWriteScreen "F0405", 6, 36
				transmit
				IF pages > 0 THEN	
					FOR i = 1 to pages
						PF8
					NEXT
				END IF
				EMSetCursor row, 30
				transmit
				call navigate_to_PRISM_screen("CPID")
				EMWriteScreen "B", 3, 29
				EMWriteScreen CP_MCI, 4, 7
				transmit
			END IF

		END IF	
		row = row + 1

		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO

		IF row = 19 THEN 'paginate
			PF8
			row = 8
			pages = pages + 1
		END IF
	LOOP UNTIL end_of_data_check = "*** End of Data ***"

	If placeholder_string = "" THEN
		msgbox "EVRs not sent for CP as there are no active employers."
		cp_evr_sent = FALSE
	END IF	
END IF

'Send EVRS to all NCP's active employers.

IF NCP_EVR_check = checked THEN
	call navigate_to_PRISM_screen("NCID")
	EMWriteScreen "B", 3, 29
	EMWriteScreen NCP_MCI, 8, 15
	transmit
	
	pages = 0
	row = 8
	placeholder_string = ""
	DO
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO

		EMReadScreen employer_code, 10, row, 30
		EMReadScreen emp_end_date, 8, row, 16

		IF trim(emp_end_date) = "" THEN
			IF Instr(placeholder_string, employer_code) = 0 THEN
				placeholder_string = placeholder_string & "~~~" & employer_code
			
				call navigate_to_PRISM_screen("DORD")
				EMWriteScreen "C", 3, 29
				transmit
				EMWriteScreen "A", 3, 29
				EMWriteScreen NCP_MCI, 4, 15
				EMWriteScreen "  ", 4, 26
				EMWriteScreen "F0405", 6, 36
				transmit
				IF pages > 0 THEN	
					FOR i = 1 to pages
						PF8
					NEXT
				END IF
				EMSetCursor row, 30
				transmit
				call navigate_to_PRISM_screen("NCID")
				EMWriteScreen "B", 3, 29
				EMWriteScreen NCP_MCI, 4, 7
				transmit
			END IF

		END IF	
		row = row + 1

		IF row = 19 THEN 'paginate
			PF8
			row = 8
			pages = pages + 1
		END IF	
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO
	LOOP UNTIL end_of_data_check = "*** End of Data ***"

	If placeholder_string = "" THEN
		msgbox "EVRs not sent for NCP as there are no active employers."
		NCP_EVR_Sent = FALSE
	END IF
END IF


'Send request to Federal Parent Locate Services (FPLS) for NCP information.

IF NCP_NCMR_check = checked THEN

	call navigate_to_PRISM_screen("NCMR")
	
	'First, make sure we have the right MCI Number displayed.
	EMWriteScreen "C", 3, 29
	EMWriteScreen NCP_mci, 4, 7   
	transmit
	EMReadScreen ssn1, 3, 6, 7
	EMReadScreen ssn2, 2, 6, 11
	EMReadScreen ssn3, 4, 6, 14
	
	'Then create the FPLS request.
	EMWriteScreen "A", 3, 29
	EMWriteScreen "FPL", 4, 27
	EMWriteScreen ssn1, 11, 7
	EMWriteScreen ssn2, 11, 11
	EMWriteScreen ssn3, 11, 14
	EMWriteScreen NCP_ln, 9, 8
	EMWriteScreen NCP_fn, 9, 33
	EMSetCursor 4, 36
	EMSendKey "<EraseEof>"
	transmit

	'Check for error messages and produce a message box to alert user if an error is rec'd.
	EMReadScreen NCP_message, 74, 24, 2
	IF InStr(NCP_message, "added successfully") = 0 THEN
		Msgbox "Unable to send a FPLS NCMR for NCP " & NCP_fn & " " & NCP_ln & ".  Error message rec'd: " & vbCr & vbCr & NCP_message 
		NCP_NCMR_SENT = FALSE
	END IF
END IF

'Send out request to FPLS for each child's MCI.
IF CH_CHMR_check = checked THEN
	call navigate_to_PRISM_screen("CHDE")
	EMWriteScreen "B", 3, 29
	transmit

	'First, make sure we have the right case displayed.
	EMWriteScreen Left(PRISM_case_number, 10), 20, 8
	EMWriteScreen Right(PRISM_case_number, 2), 20, 19
	transmit

	row = 8
	ch_placeholder_string = ""
	DO
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO

		EMReadScreen child_mci, 10, row, 67
		EMReadScreen child_name, 30, row, 3
		child_name = trim(child_name)	
		'A placeholder variable is used to keep track of requests that have been sent out, referencing the children's unique MCI number.  
		'Only sending CHMR requests if it hasn't been done yet for this child/MCI.
		IF Instr(ch_placeholder_string, child_mci) = 0 THEN
				ch_placeholder_string = ch_placeholder_string & "~~~" & child_mci
			
				call navigate_to_PRISM_screen("CHMR")
				EMWriteScreen "C", 3, 29
				transmit
				EMWriteScreen "A", 3, 29
				EMWriteScreen child_mci, 4, 7
				EMWriteScreen "FPL", 4, 27
				EMSetCursor 4, 36
				EMSendKey "<EraseEof>"
				transmit
				
				'Check for error messages and produce a message box to alert user if an error is rec'd.
				EMReadScreen message, 74, 24, 2

				' NOTE: In the test region, PRISM did not produce an error message when a duplicate request is sent.  However, error fired when 
				' sending out a FPL request on an inactive child. This behavior differs from NCP requests.

				IF InStr(message, "added successfully") = 0 THEN
					Msgbox "Unable to send a FPLS CHMR for " & child_name & ".  Error message rec'd: " & vbCr & vbCr & message 
					CHMR_SENT = FALSE
					children_not_sent = children_not_sent & " " & child_name & "-"
				ELSE
					children_sent = children_sent & " " & child_name & "-"	
				END IF
		END IF		
		call navigate_to_PRISM_screen("CHDE")
		EMWriteScreen "B", 3, 29
		transmit
		EMWriteScreen PRISM_left, 20, 8
		EMWriteScreen PRISM_right, 20, 19
		transmit
	
		row = row + 1
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO
		IF row = 19 THEN 'paginate
			PF8
			row = 8
			pages = pages + 1
		END IF
	LOOP UNTIL end_of_data_check = "*** End of Data ***"
END IF


'Update pending Notice of Review document with mod worker information.
IF update_mod_worker = checked THEN
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "B", 3, 29
	transmit

	'First, make sure we have the right case displayed.
	EMWriteScreen PRISM_left, 20, 26
	EMWriteScreen PRISM_right, 20, 37


	EMWriteScreen "PND", 20, 48
	transmit

	row = 5
	dord_placeholder_string = ""

	DO
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO
		 
		EMReadScreen stat, 3, row, 22
		EMReadScreen doc_name, 30, row, 57
		

		IF InStr(doc_name, "REVIEW SUPPORT") > 0 AND stat = "PND" THEN			
				EMWriteScreen "S", row, 5
				transmit
				
				PF14
				EMWriteScreen "U", 20, 14
				transmit
				EMWriteScreen "S", 7, 5
				EMWriteScreen "S", 8, 5
				EMWriteScreen "S", 12, 5
				transmit
				
		
				IF mod_worker = "Theresa Hogan" THEN
					EMWriteScreen "Theresa Hogan          ", 16, 15
					transmit
					EMWriteScreen "Expedited Process Specialist", 16, 15
					transmit
					EMWriteScreen "763-323-6058", 16, 15
					transmit
				ELSEIF mod_worker = "Terri Spence-Garski" THEN
					EMWriteScreen "Terri Spence-Garski", 16, 15
					transmit
					EMWriteScreen "Expedited Process Specialist", 16, 15
					transmit
					EMWriteScreen "763-323-6056", 16, 15				
					transmit
				ELSEIF mod_worker = "Kelly Rein" THEN
					EMWriteScreen "Kelly Rein                  ", 16, 15
					transmit
					EMWriteScreen "Expedited Process Specialist", 16, 15
					transmit
					EMWriteScreen "763-323-6059", 16, 15					 
					transmit
				ELSEIF mod_worker = "Karla Wangrud" THEN
					EMWriteScreen "Karla Wangrud            ", 16, 15
					transmit
					EMWriteScreen "Expedited Process Specialist", 16, 15
					transmit
					EMWriteScreen "763-422-7346", 16, 15					 
					transmit
				END IF		
			
				PF3 
				EMWriteScreen "M", 3, 29
				transmit

				Dialog LH_dialog  'name of dialog
			  	IF buttonpressed = 0 then script_end_procedure("The script has ended.")		'Cancel
				
				call navigate_to_PRISM_screen("DORD")
				EMWriteScreen "B", 3, 29
				transmit
	
				'Make sure we have the right case displayed.
				EMWriteScreen PRISM_left, 20, 26
				EMWriteScreen PRISM_right, 20, 37
	

				EMWriteScreen "PND", 20, 48
				transmit
			
		END IF 
		
		row = row + 1
		EMReadScreen end_of_data_check, 19, row, 28
		IF end_of_data_check = "*** End of Data ***" THEN EXIT DO
		IF row = 18 THEN 'paginate
			PF8
			row = 5
			pages = pages + 1
		END IF
	LOOP UNTIL end_of_data_check = "*** End of Data ***"
	

END IF

'Creating followup worklist
If CAWD_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")

	'First, make sure we have the right case displayed.
	EMWriteScreen PRISM_left, 20, 8
	EMWriteScreen PRISM_right, 20, 19
	transmit
	
	'Then, add worklist
	PF5
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** Special Services assessment received from the parties?", 10, 4
	EMWriteScreen dateadd("d", date, 7), 17, 21
	transmit
End if

'Going to CAAD, adding a new note
call navigate_to_PRISM_screen("CAAD")
'First, make sure we are have the right case displayed.
EMWriteScreen "A", 8, 5
EMWriteScreen PRISM_left, 20, 8
EMWriteScreen PRISM_right, 20, 19																
transmit
PF5

'Make sure we are on the right page.  If not, display a message to the user.
EMReadScreen case_activity_detail, 20, 2, 29

If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")
	'Then, add case note
	EMWriteScreen "M2123", 4, 54

	'Setting cursor in write area and writing note details
	EMSetCursor 16, 4
	call write_variable_in_CAAD("* The following documents sent:")
	If CP_Special_assessment_check = checked then call write_variable_in_CAAD("    * Coverletter and Special Services Assessment to CP")
	If NCP_Special_assessment_check = checked then call write_variable_in_CAAD("    * Coverletter and Special Services Assessment to NCP")
	If child_care_verification_check = checked then call write_variable_in_CAAD("    * Child care verification sent to CP")
	if CH_CHMR_check = checked then 
		IF CHMR_SENT = FALSE THEN
			call write_variable_in_CAAD("* Error sending manual FPLS locate sent for child(ren):" & children_not_sent)
		ELSE
			call write_variable_in_CAAD("* Manual FPLS locate sent for child(ren):" & children_sent)
		END IF
	END IF
	if NCP_NCMR_check = checked then
		IF NCP_NCMR_SENT = FALSE THEN
			 call write_variable_in_CAAD("* Error sending manual FPLS locate for NCP")

		ELSE
			 call write_variable_in_CAAD("* Manual FPLS locate sent for NCP")	
		END IF
	END IF	
	
	IF NCP_EVR_check = checked THEN
		IF NCP_EVR_SENT = FALSE THEN 
			CALL write_variable_in_CAAD(" * EVR not sent for NCP as no active employers")
		ELSE
			CALL write_variable_in_CAAD(" * EVR requests sent to NCP's employer(s)")
		END IF
	END IF

	IF CP_EVR_check = checked THEN 
		IF CP_EVR_SENT = FALSE THEN
			CALL write_variable_in_CAAD(" * EVR not sent for CP as no active employers")
		ELSE
			CALL write_variable_in_CAAD(" * EVR requests sent to CP's employer(s)")
		END IF
	END IF

	if CAWD_check = checked then call write_variable_in_CAAD("* Followup worklist created")
	call write_variable_in_CAAD(CAAD_note_text)
	call write_variable_in_CAAD("---")
	call write_variable_in_CAAD("* Documents to be returned by " & dateadd("d", date, 7) & ".")
	call write_variable_in_CAAD("---")
	call write_variable_in_CAAD(worker_signature)

script_end_procedure("The script is now ending.  Please save your CAAD note.")
