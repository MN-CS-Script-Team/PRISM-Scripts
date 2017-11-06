'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "mets-referral-attempted.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/28/2017", "Initial version.", "Kelly Hiestand, Wright County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'THE DIALOG-------------------------------------------------------------------------------------------------------------
BeginDialog mets_referral_dialog, 0, 0, 316, 190, "METS/MAO Referral Attempted"
  EditBox 80, 5, 95, 15, prism_case_number
  EditBox 75, 25, 75, 15, mets_case_number
  EditBox 245, 25, 60, 15, date_ref_attempt
  EditBox 55, 45, 95, 15, mets_county
  EditBox 115, 65, 135, 15, children
  EditBox 115, 85, 135, 15, error_msg
  EditBox 50, 110, 200, 15, other_notes
  DropListBox 110, 140, 140, 15, "Select one..."+chr(9)+"Create Worklist"+chr(9)+"Generate Email"+chr(9)+"Do not create a Worklist or Email", action_dropdown
  EditBox 70, 165, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 165, 50, 15
    CancelButton 260, 165, 50, 15
  Text 160, 30, 80, 10, "Date Referral Attempted:"
  Text 5, 170, 60, 10, "Worker Signature:"
  Text 5, 30, 70, 10, "METS Case Number:"
  Text 5, 135, 100, 25, "Do you want to create a worklist or generate an email to the worker with this info?"
  Text 5, 90, 105, 10, "Error Message Rec'd in METS:"
  Text 5, 10, 75, 10, "PRISM Case Number:"
  Text 5, 70, 105, 10, "Referral Attempted for Child/ren:"
  Text 5, 50, 50, 10, "METS County:"
  Text 5, 115, 45, 10, "Other Notes:"
EndDialog


'Connecting to Bluezone
EMConnect ""			

'MAKES SURE YOU ARE NOT PASSWORDED OUT
CALL check_for_PRISM(True)

'AUTO POPULATES PRISM CASE NUMBER INTO DIALOG
CALL PRISM_case_number_finder(PRISM_case_number)

'MAKES THINGS MANDATORY
DO
	err_msg = ""
	Dialog mets_referral_dialog
	IF buttonpressed = 0 THEN stopscript
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF mets_case_number = "" THEN err_msg = err_msg & vbNewline & "You must enter an 8 digit METS case number!"
	IF len(mets_case_number)<> 8 THEN err_msg = err_msg & vbNewline & "The METS case number must be 8 digits!"
	IF action_dropdown = "Select One..." THEN err_msg = err_msg & vbNewline & "You must select how to notify the worker!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'NAVIGATES TO CAAD
CALL navigate_to_PRISM_screen("CAAD")

'ENTERING CASE NUMBER
CALL enter_PRISM_case_number(PRISM_case_number, 20, 8)

'ADDS NEW CAAD NOTE WITH FREE CAAD CODE
PF5
EMWritescreen "FREE", 4, 54

'CLEANING UP THE LANGUAGE FOR THE CAAD NOTE, THE TENSE WAS NOT GRAMATICALLY CORRECT
IF action_dropdown = "Create Worklist" THEN action_dropdown = "Created Worklist"
IF action_dropdown = "Generate Email" THEN action_dropdown = "Email"
IF action_dropdown = "Do not create a Worklist or Email" THEN action_dropdown = "Did not notify worker"

'SETS THE CURSOR
EMSetCursor 16, 4

'WRITES THE CAAD NOTE
CALL write_variable_in_CAAD("METS/MAO Referral Attempted - Failed")
CALL write_bullet_and_variable_in_CAAD("METS Case Number", mets_case_number)
CALL write_bullet_and_variable_in_CAAD("METS County", mets_county)
CALL write_bullet_and_variable_in_CAAD("Date Referral Attempted", date_ref_attempt)
CALL write_bullet_and_variable_in_CAAD("Referral Attempted for child/ren", children)
CALL write_bullet_and_variable_in_CAAD("Error message rec'd in METS", error_msg)
CALL write_bullet_and_variable_in_CAAD("Other Notes", other_notes)
CALL write_bullet_and_variable_in_CAAD("Notified Worker Via", action_dropdown)
CALL write_variable_in_CAAD(worker_signature)
transmit


'ADDS A WORKLIST IF THE DROPDOWN TO ADD ONE IS SELECTED
IF action_dropdown = "Created Worklist" THEN
	CALL navigate_to_PRISM_screen("CAWT")
	PF5
	EMWritescreen "FREE", 4, 37
	
	'SETS THE CURSOR AND STARTS THE WORKLIST
	EMWriteScreen "METS/MA Referral Attempted - Failed", 10, 4
	EMWriteScreen "Referral Attempted: " & date_ref_attempt, 11, 4
	EMWriteScreen "METS Case Number: " & mets_case_number, 11, 40
	EMWriteScreen "Error message rec'd, MAO referral unable to be sent at this time", 12, 4
	EMWriteScreen worker_signature, 13, 4
	EMWritescreen "10", 17, 52
	transmit
END IF

'CREATES AN EMAIL IF THE DROPDOWN TO SEND AN EMAIL IS SELECTED
IF action_dropdown = "Email" THEN
	CALL create_outlook_email("", "", "Unable to sent MAO referral from METS", "PRISM CASE NUMBER: " & prism_case_number & vbcr & "METS CASE NUMBER: " & mets_case_number & vbcr & "METS COUNTY: " & mets_county & vbcr & "The MAO referral was attempted to be sent from METS, but failed because of known issues in the METS system. A FREE CAAD note was also created on this case.", "", FALSE)
End IF

script_end_procedure("")
