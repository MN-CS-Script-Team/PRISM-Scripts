'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "unreimbursed-uninsured-returned-docs.vbs"
start_time = timer
STATS_counter = 1
'STATS_manualtime = 
STATS_denomination = "C"

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
call changelog_update("03/28/2017", "You can now add the County Name for the Affidavit of Service.", "Gretchen Thornbrugh, Dakota Co.")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'first dialog 
BeginDialog UnUn1_Dialog, 0, 0, 276, 135, "Unreimbursed Uninsured Docs Received"
  EditBox 70, 5, 80, 15, PRISM_case_number
  DropListBox 195, 65, 60, 45, "Select One..."+chr(9)+"CPP"+chr(9)+"NCP", person_droplistbox
  ButtonGroup ButtonPressed
    OkButton 155, 110, 50, 15
    CancelButton 215, 110, 50, 15
  Text 15, 10, 50, 10, "Case Number"
  Text 15, 30, 255, 25, "This script will gernerate DORD DOCS Notice of Intent to Enforce UN/UN and Affidavit of Service for the collection of Unreimbursed and Uninsured Medical and Dental Expenses as requested by CP or NCP."
  Text 5, 70, 190, 10, "Select who returned the Unreimbursed/Uninsured forms."
EndDialog
'end first dialog 

'Connecting to BlueZone
EMConnect ""

'brings me to the CAPS screen
CALL navigate_to_PRISM_screen ("CAPS")

'check for prism (password out)before continuing
CALL check_for_PRISM(true)

'this auto fills prism case number in dialog
CALL PRISM_case_number_finder(PRISM_case_number)

'THE LOOP--------------------------------------
'adding a loop

Do
	err_msg = ""
	Dialog UnUn1_Dialog 'Shows name of dialog
	IF buttonpressed = 0 then stopscript		'Cancel
	IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed"
	IF person_droplistbox = "Select One..." THEN err_msg = err_msg & vbNewline & "Select who returned the documents."
	IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF

LOOP UNTIL err_msg = ""


'dialog if cp is requesting collection of un/un
IF person_droplistbox = "CPP" THEN
	
	BeginDialog Cp_requested_Dialog, 0, 0, 216, 200, "Completed Documents received from Requesting Party (CP)"
  	  EditBox 80, 5, 60, 15, amount
  	  CheckBox 10, 30, 195, 10, "Check to add CAAD note of documents received from CP", CPCAAD_checkbox
  	  EditBox 85, 80, 50, 15, worker_signature
  	  CheckBox 15, 110, 170, 10, "Check to add information to JUDE, OBBD, NCOD", jude_checkbox
  	  CheckBox 15, 130, 200, 10, "Check to send documents to Non Requesting Party (NCP)", ncp_documents_checkbox
 	  ButtonGroup ButtonPressed
    	    OkButton 95, 180, 50, 15
    	    CancelButton 155, 180, 50, 15
  	  Text 10, 10, 65, 10, "Amount Requested:"
  	  Text 35, 45, 115, 10, "Affidavit of Health Care Expenses"
  	  Text 35, 55, 145, 10, "Notice to Collect UN Med Exp Req Party"
  	  Text 35, 65, 110, 10, "Copies of bill, receipts, EOB's"
  	  Text 20, 85, 60, 10, "Worker Signature"
  	  Text 35, 145, 115, 10, "Notice of Intent to Enforce UN/UN"
  	  Text 35, 155, 65, 10, "Affidavit of Service"
	EndDialog

Do
	err_msg = ""
	Dialog Cp_requested_Dialog 'Shows name of dialog
	IF buttonpressed = 0 then stopscript		'Cancel
	IF CPCAAD_checkbox = 0 AND jude_checkbox = 0 AND ncp_documents_checkbox = 0 THEN err_msg = err_msg & vbNewline & "Please select at least one checkbox."
	IF amount = "" THEN err_msg = err_msg & vbNewline & "The UnUn amount to be collected must be completed." 
	IF CPCAAD_checkbox =1 AND worker_signature = "" THEN err_msg = err_msg & vbNewline & "Please sign your CAAD Note."
	IF err_msg <> "" THEN 
		MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF

LOOP UNTIL err_msg = ""


'ADDS CAAD NOTE
	IF CPCAAD_checkbox = 1 THEN
		CALL navigate_to_PRISM_screen ("CAAD")																					
		PF5
		EMWriteScreen "A", 3, 29
		EMWriteScreen "free", 4, 54
		EMSetCursor 16, 4

'this will add information to the CAAD note of what emc docs sent 
		CALL write_variable_in_CAAD ("CP returned Affidavit of Health Care Expenses, Notice to Collect UN MED   Exp Req Party, and Copies of bills, receipts, EOB's.")
		CALL write_variable_in_CAAD ("Amount requested $" & amount)
		CALL write_variable_in_CAAD(worker_signature)
		transmit
		PF3
	END IF
'END IF

	IF ncp_documents_checkbox = 1 THEN

	BeginDialog DATE_SERVED_dialog, 0, 0, 161, 95, "DATE SERVED"
  	  EditBox 50, 5, 50, 15, date_served
  	  EditBox 65, 30, 65, 15, county_name
  	  CheckBox 10, 55, 125, 10, "check if address is CONFIDENTIAL", confidential_checkbox
  	  ButtonGroup ButtonPressed
    	    OkButton 50, 75, 50, 15
          CancelButton 105, 75, 50, 15
  	  Text 10, 10, 40, 10, "Served on:"
        Text 10, 35, 50, 10, "County Name:"
	EndDialog

	
'dialog box for date on aff of service
Do
	err_msg = ""
	Dialog DATE_SERVED_dialog
	IF buttonpressed = 0 then stopscript
	IF date_served = "" THEN err_msg = err_msg & vbNewline & "Please enter date you are sending Affidavit of Service."
	IF county_name = "" THEN err_msg = err_msg & vbNewline & "Please enter County Name for the Affidavit of Service."
	IF err_msg <> "" THEN 
		MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF
Loop until err_msg = ""

'creates notice of intent to enforce
		CALL navigate_to_PRISM_screen ("DORD")
		EMWriteScreen "C", 3, 29
		transmit

		EMWriteScreen "A", 3, 29
		EMWriteScreen "F0949", 6, 36
		EMWriteScreen "ncp", 11, 51
		transmit
		PF14
		PF8
		PF8	

		EMWriteScreen "S", 11, 5
		transmit 
	
		EMWriteScreen amount, 16, 15
		transmit
		PF3
		EMWriteScreen "m", 3, 29
		transmit

	END IF

'DORD aff of service
	IF ncp_documents_checkbox = 1 AND confidential_checkbox = 0 THEN
		CALL navigate_to_PRISM_screen ("DORD")
		EMWriteScreen "C", 3, 29
		transmit

		EMWriteScreen "A", 3, 29
		EMWriteScreen "F0016", 6, 36
		EMWriteScreen "ncp", 11, 51
		transmit
'shift f2, to get to user lables
		PF14
		EMWriteScreen "u", 20, 14
		transmit
		PF8
		PF8
		EMWriteScreen "s", 15, 5
		EMWriteScreen "s", 16, 5
		EMWriteScreen "s", 17, 5
		transmit
		EMWriteScreen "Notice of Intent to Enforce Unreimbursed and/or Uninsured", 16, 15
		transmit
		EMWriteScreen "Medical/Dental Expenses", 16, 15
		transmit
		EMWriteScreen date_served, 16, 15
		transmit
		PF8
		EMWriteScreen "s", 8, 5
		EMWriteScreen "s", 10, 5
		transmit
		EMWriteScreen "N", 16, 15
		transmit
		EmWriteScreen county_name, 16, 15
		transmit
		PF3
		EMWriteScreen "M", 3, 29
		transmit

'''need to select legal heading
		BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  	 	  ButtonGroup ButtonPressed
    	    	    OkButton 60, 75, 50, 15
    	    	    CancelButton 115, 75, 50, 15
  		  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
 		  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  		  Text 5, 40, 55, 10, "2. Press ENTER"
  		  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
		EndDialog

		Dialog LH_dialog  'name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
			
	END IF


	IF ncp_documents_checkbox = 1 AND confidential_checkbox = 1 THEN
		CALL navigate_to_PRISM_screen ("DORD")
		EMWriteScreen "C", 3, 29
		transmit

		EMWriteScreen "A", 3, 29
		EMWriteScreen "F0016", 6, 36
		EMWriteScreen "ncp", 11, 51
		transmit
'shift f2, to get to user lables
		PF14
		EMWriteScreen "u", 20, 14
		transmit
		PF8
		PF8
		EMWriteScreen "s", 15, 5
		EMWriteScreen "s", 16, 5
		EMWriteScreen "s", 17, 5
		transmit
		EMWriteScreen "Notice of Intent to Enforce Unreimbursed and/or Uninsured", 16, 15
		transmit
		EMWriteScreen "Medical/Dental Expenses", 16, 15
		transmit
		EMWriteScreen date_served, 16, 15
		transmit
		PF8
		EMWriteScreen "s", 8, 5
		EMWriteScreen "s", 10, 5
		transmit
		EMWriteScreen "Y", 16, 15
		transmit
		EMWriteScreen county_name, 16, 15
		transmit
		PF3
		EMWriteScreen "M", 3, 29
		transmit

'''need to select legal heading
		BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  	  	  ButtonGroup ButtonPressed
    	    	    OkButton 60, 75, 50, 15
    	    	    CancelButton 115, 75, 50, 15
  	  	  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  	  	  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  	  	  Text 5, 40, 55, 10, "2. Press ENTER"
  	  	  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
		EndDialog

		Dialog LH_dialog  'name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
			
	END IF
'END IF


	IF jude_checkbox = 1 THEN
	'CP Name											
		call navigate_to_PRISM_screen("CPDE")
		EMWriteScreen CP_MCI, 4, 7
		EMReadScreen CP_F, 12, 8, 34
		EMReadScreen CP_M, 12, 8, 56
		EMReadScreen CP_L, 17, 8, 8

		CP_name = fix_read_data(CP_F) & " " & fix_read_data(CP_M) & " " & fix_read_data(CP_L)	
		CP_name = trim(CP_Name)


		CALL navigate_to_PRISM_screen ("SUOD")
		EMWriteScreen "B", 3, 29
		transmit

		BeginDialog PRISM_INFO_Dialog, 0, 0, 266, 185, "Info needed to add Un/Un to PRISM"
  	  	  EditBox 85, 25, 25, 15, CO_Seq
  	  	  EditBox 50, 45, 50, 15, From_date
  	 	  EditBox 130, 45, 50, 15, To_date
  	        EditBox 50, 70, 200, 15, CP_name
  	 	  EditBox 65, 110, 40, 15, eff_date
  		  EditBox 55, 135, 50, 15, beg_date
  	 	  ButtonGroup ButtonPressed
    	    	    OkButton 145, 160, 50, 15
    	    	    CancelButton 205, 160, 50, 15
  		  Text 101, 10, 65, 10, "JUDE Information"
  		  Text 10, 30, 70, 10, "Court Order Seq Nbr:"
  		  Text 120, 30, 35, 10, "format 01"
  		  Text 10, 50, 40, 10, "Date From:"
  		  Text 110, 50, 15, 10, "To:"
  		  Text 190, 50, 50, 10, "xx/xx/xxxx"
  		  Text 10, 75, 40, 10, "In Favor of:"
  		  Text 101, 95, 65, 10, "NCOD Information"
  		  Text 10, 115, 55, 10, "Effective Date:"
  		  Text 130, 115, 50, 10, "xx/xxxx"
  		  Text 10, 140, 40, 10, "Begin Date:"
  		  Text 120, 140, 50, 10, "xx/xx/xxxx"
		EndDialog


Do
	err_msg = ""
	Dialog PRISM_INFO_Dialog
		IF buttonpressed = 0 then stopscript
		IF Co_Seq = "" THEN err_msg = err_msg & vbNewline & "Please enter the Court order sequence number."
		IF From_date = "" THEN err_msg = err_msg & vbNewline & "Please enter FROM date."
		IF To_date = "" THEN err_msg = err_msg & vbNewline & "Please enter TO date."
		IF CP_name = "" THEN err_msg = err_msg & vbNewline & "Please enter the CP's name."
		IF eff_date = "" THEN err_msg = err_msg & vbNewline & "Please enter the effective date."
		IF beg_date = "" THEN err_msg = err_msg & vbNewline & "Please enter the begin date."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF
Loop until err_msg = ""

	'adding jude info
		CALL navigate_to_PRISM_screen ("JUDE")
		EMWriteScreen "C", 3, 29
		transmit
		EMWriteScreen "A", 3, 29
		EMWriteScreen Co_Seq, 4, 34
		EMWriteScreen "JME", 10, 6
		EMWriteScreen From_date, 10, 17
		EMWriteScreen To_date, 10, 31
		EMWriteScreen CP_name, 13, 16
		EMWriteScreen amount, 14, 17
		EMWriteScreen "JOL", 15, 20
		PF11
		EMWriteScreen "un/un expenses requested by cp", 12, 3
		transmit

	'checking bottom screen for jol success
		EMReadScreen jol_success, 18, 24, 33
		IF jol_success <> "added successfully" THEN 
			script_end_procedure ("Jude information was not added correctly, please reneter information.  Script Ended.")
		END IF

	'reading judgment sequence number to add to ncod
		EMReadScreen jdgmt_number, 2, 4, 52
		
	'adding ncod info
		CALL navigate_to_PRISM_screen ("NCOD")
		EMWriteScreen "C", 3, 29
		transmit
		EMWriteScreen "A", 3, 29
		EMWriteScreen "JME", 4, 34
		EMWriteScreen "  ", 4, 053
		EMWriteScreen eff_date, 9, 59 
		EMWriteScreen "npa", 12, 10
		EMWriteScreen Co_Seq, 11, 62
		EMWriteScreen "n", 13, 12
		EMWriteScreen Co_Seq, 12, 55
		EMWriteScreen jdgmt_number, 12, 74
		EMWriteScreen "y", 18, 57
		EMWriteScreen beg_date, 14, 68 
		transmit
		'transmit
	
	'reading ncod success
		EMReadScreen ncod_success, 18 , 24, 34
		IF ncod_success <> "added successfully" THEN 
			ncod_message = Msgbox ("NCOD information was not added correctly, please correct error and click OK to continue. click CANCEL to end script.", VbOKCancel)
			If ncod_message = vbCancel then stopscript
	END IF

	'adding obbd info
	CALL navigate_to_PRISM_screen ("OBBD")
	EMWriteScreen "M", 3, 29
	EMWriteScreen "           ", 18, 15
	EMWriteScreen amount, 18, 15
	PF11
	EMWriteScreen "added un/un expenses. " & worker_signature, 18, 25
	EMWriteScreen "n", 17, 72 
	transmit

	'reading modified sucess
	EMReadScreen obbd_success, 13 , 24, 68
		IF obbd_success <> "modified succ" THEN 
			Msgbox "OBBD information was not added correctly, please reneter information.  Script Ended."
			StopScript
		END IF

CALL navigate_to_PRISM_screen ("NCOL")

	END IF
END IF

	
'*****************************dialog if ncp is requesting collection of un/un from the cp********************************

IF person_droplistbox = "NCP" THEN

	BeginDialog Ncp_requested_Dialog, 0, 0, 216, 200, "Completed Documents received from Requesting Party (NCP)"
  	  EditBox 80, 5, 60, 15, amount
 	  CheckBox 10, 30, 195, 10, "Check to add CAAD note of documents received from NCP", NCPCAAD_checkbox
	  EditBox 85, 80, 50, 15, worker_signature
	  CheckBox 15, 110, 75, 10, "Check to add CPOD", cpod_checkbox
	  CheckBox 15, 130, 190, 10, "Check to send documents to Non Requesting Party (CP)", cp_documents_checkbox
	  ButtonGroup ButtonPressed
    	    OkButton 95, 180, 50, 15
    	    CancelButton 155, 180, 50, 15
	  Text 10, 10, 65, 10, "Amount Requested:"
	  Text 35, 45, 115, 10, "Affidavit of Health Care Expenses"
	  Text 35, 55, 145, 10, "Notice to Collect UN Med Exp Req Party"
	  Text 35, 65, 110, 10, "Copies of bill, receipts, EOB's"
	  Text 20, 85, 60, 10, "Worker Signature"
	  Text 35, 145, 115, 10, "Notice of Intent to Enforce UN/UN"
 	 Text 35, 155, 65, 10, "Affidavit of Service"
	EndDialog


Do
	err_msg = ""
	Dialog Ncp_requested_Dialog 'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF NCPCAAD_checkbox = 0 AND cpod_checkbox = 0 AND cp_documents_checkbox = 0 THEN err_msg = err_msg & vbNewline & "Please select at least one checkbox."
		IF amount = "" THEN err_msg = err_msg & vbNewline & "The UnUn amount to be collected must be completed." 
		IF NCPCAAD_checkbox =1 AND worker_signature = "" THEN err_msg = err_msg & vbNewline & "Please sign your CAAD Note."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'ADDS CAAD NOTE
	IF NCPCAAD_checkbox = 1 THEN
		CALL navigate_to_PRISM_screen ("CAAD")																					
		PF5
		EMWriteScreen "A", 3, 29
		EMWriteScreen "free", 4, 54
		EMSetCursor 16, 4

'this will add information to the CAAD note of what emc docs sent 
		CALL write_variable_in_CAAD ("NCP returned Affidavit of Health Care Expenses, Notice to Collect UN MED   Exp Req Party, and Copies of bills, receipts, EOB's.")
		CALL write_variable_in_CAAD ("Amount requested $" & amount)
		CALL write_variable_in_CAAD(worker_signature)
		transmit
		PF3
	END IF

	IF cp_documents_checkbox = 1 THEN

	BeginDialog DATE_SERVED_dialog, 0, 0, 161, 95, "DATE SERVED"
  	  EditBox 50, 5, 50, 15, date_served
  	  EditBox 65, 30, 65, 15, county_name
  	  CheckBox 10, 55, 125, 10, "check if address is CONFIDENTIAL", confidential_checkbox
  	  ButtonGroup ButtonPressed
    	    OkButton 50, 75, 50, 15
    	    CancelButton 105, 75, 50, 15
  	  Text 10, 10, 40, 10, "Served on:"
 	  Text 10, 35, 50, 10, "County Name:"
	EndDialog

'dialog box for date on aff of service

Do
	err_msg = ""
	Dialog DATE_SERVED_dialog
	IF buttonpressed = 0 then stopscript
	IF date_served = "" THEN err_msg = err_msg & vbNewline & "Please enter date you are sending Affidavit of Service."
	IF county_name = "" THEN err_msg = err_msg & vbNewline & "Please enter County Name for the Affidavit of Service."	
	IF err_msg <> "" THEN 
		MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF

Loop until err_msg = ""

'creates notice of intent to enforce
	CALL navigate_to_PRISM_screen ("DORD")
		EMWriteScreen "C", 3, 29
		transmit

		EMWriteScreen "A", 3, 29
		EMWriteScreen "F0949", 6, 36
		EMWriteScreen "cpp", 11, 51
		transmit
		PF14
		PF8
		PF8	

		EMWriteScreen "S", 11, 5
		transmit 
	
		EMWriteScreen amount, 16, 15
		transmit
		PF3
		EMWriteScreen "m", 3, 29
		transmit

	END IF

'DORD aff of service
	IF cp_documents_checkbox = 1 AND confidential_checkbox = 0 THEN
		CALL navigate_to_PRISM_screen ("DORD")
		EMWriteScreen "C", 3, 29
		transmit

		EMWriteScreen "A", 3, 29
		EMWriteScreen "F0016", 6, 36
		EMWriteScreen "cpp", 11, 51
		transmit
'shift f2, to get to user lables
		PF14
		EMWriteScreen "u", 20, 14
		transmit
		PF8
		PF8
		EMWriteScreen "s", 15, 5
		EMWriteScreen "s", 16, 5
		EMWriteScreen "s", 17, 5
		transmit
		EMWriteScreen "Notice of Intent to Enforce Unreimbursed and/or Uninsured", 16, 15
		transmit
		EMWriteScreen "Medical/Dental Expenses", 16, 15
		transmit
		EMWriteScreen date_served, 16, 15
		transmit
		PF8
		EMWriteScreen "s", 8, 5
		EMWriteScreen "s", 10, 5
		transmit
		EMWriteScreen "N", 16, 15
		transmit
		EMWriteScreen county_name, 16, 15
		transmit
		PF3
		EMWriteScreen "M", 3, 29
		transmit

'''need to select legal heading
		BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
	  	  ButtonGroup ButtonPressed
 	   	    OkButton 60, 75, 50, 15
    	    	    CancelButton 115, 75, 50, 15
	 	  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
		  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
		  Text 5, 40, 55, 10, "2. Press ENTER"
 		  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
		EndDialog

		Dialog LH_dialog  'name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
			
	END IF


	IF cp_documents_checkbox = 1 AND confidential_checkbox = 1 THEN
		CALL navigate_to_PRISM_screen ("DORD")
		EMWriteScreen "C", 3, 29
		transmit

		EMWriteScreen "A", 3, 29
		EMWriteScreen "F0016", 6, 36
		EMWriteScreen "cpp", 11, 51
		transmit
'shift f2, to get to user lables
		PF14
		EMWriteScreen "u", 20, 14
		transmit
		PF8
		PF8
		EMWriteScreen "s", 15, 5
		EMWriteScreen "s", 16, 5
		EMWriteScreen "s", 17, 5
		transmit
		EMWriteScreen "Notice of Intent to Enforce Unreimbursed and/or Uninsured", 16, 15
		transmit
		EMWriteScreen "Medical/Dental Expenses", 16, 15
		transmit
		EMWriteScreen date_served, 16, 15
		transmit
		PF8
		EMWriteScreen "s", 8, 5
		EMWriteScreen "s", 10, 5
		transmit
		EMWriteScreen "Y", 16, 15
		transmit
		EMWriteScreen county_name, 16, 15
		transmit
		PF3
		EMWriteScreen "M", 3, 29
		transmit

'''need to select legal heading
		BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
		  ButtonGroup ButtonPressed
    	    	    OkButton 60, 75, 50, 15
         	    CancelButton 115, 75, 50, 15
		  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
		  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
		  Text 5, 40, 55, 10, "2. Press ENTER"
		  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
		EndDialog

		Dialog LH_dialog  'name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
	END IF



	IF cpod_checkbox = 1 THEN
		CALL navigate_to_PRISM_screen ("SUOD")
		EMWriteScreen "B", 3, 29
		transmit

		BeginDialog CPOD_Dialog, 0, 0, 176, 120, "CPOD"
  		  EditBox 85, 25, 25, 15, CO_Seq
 		  EditBox 65, 45, 40, 15, eff_date
		  EditBox 55, 65, 50, 15, beg_date
  		  ButtonGroup ButtonPressed
  		    OkButton 60, 100, 50, 15
  		    CancelButton 120, 100, 50, 15
		  Text 56, 10, 65, 10, "CPOD  Information"
		  Text 10, 30, 70, 10, "Court Order Seq Nbr:"
		  Text 120, 30, 35, 10, "format 01"
		  Text 10, 50, 55, 10, "Effective Date:"
		  Text 125, 50, 35, 10, "xx/xxxx"
		  Text 10, 70, 40, 10, "Begin Date:"
 		  Text 115, 70, 45, 10, "xx/xx/xxxx"
		EndDialog


Do
	err_msg = ""
	Dialog CPOD_Dialog
	IF buttonpressed = 0 then stopscript
	IF Co_Seq = "" THEN err_msg = err_msg & vbNewline & "Please enter the Court order sequence number."
	IF eff_date = "" THEN err_msg = err_msg & vbNewline & "Please enter the effective date."
	IF beg_date = "" THEN err_msg = err_msg & vbNewline & "Please enter the begin date."
	IF err_msg <> "" THEN 
		MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF

Loop until err_msg = ""


'add information on cpod
		CALL navigate_to_PRISM_screen ("CPOD")
		EMWriteScreen "C", 3, 29
		transmit
		EMWriteScreen "A", 3, 29
		EMSetCursor 4, 53
		EMWriteScreen "  ", 4, 53
		EMWriteScreen "JME", 4, 34
		EMWriteScreen "DIR", 9, 35
		EMWriteScreen eff_date, 9, 59
		EMWriteScreen "MDN", 12, 10 
		EMWriteScreen "N", 13, 12 
		EMWriteScreen Co_Seq, 12, 55 
		EMWriteScreen beg_date, 14, 68
		EMWriteScreen "D", 18, 57
		transmit
	
	
		EMReadScreen cpod_success, 18 , 24, 33
			IF cpod_success <> "added successfully" THEN 
				script_end_procedure ("CPOD information was not added correctly, please reneter information.  Script Ended.")
			END IF
	
'add information on obbd
		CALL navigate_to_PRISM_screen ("OBBD")
		EMWriteScreen "M", 3, 29
		EMSetCursor 18, 15
		EMWriteScreen "            ", 18, 15
		EMWriteScreen amount, 18, 15
		PF11
		EMWriteScreen "added un/un expenses. " & worker_signature, 18, 25
		EMWriteScreen "n", 17, 72 
		transmit

	'reading modified success 
		EMReadScreen obbd_success, 13 , 24, 66
			IF obbd_success <> "modified succ" THEN 
				script_end_procedure ("OBBD information was not added correctly, please reneter information.  Script Ended.")
			END IF

		CALL navigate_to_PRISM_screen ("CPOL")
	END IF		
END IF	
script_end_procedure("")

