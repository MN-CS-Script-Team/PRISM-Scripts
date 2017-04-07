'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "maintaining-county.vbs"


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
call changelog_update ("03/09/2017", "Removed automatic transmit so user can save the CAAD note themselves.", "Kelly Hiestand, Wright County")
call changelog_update ("01/18/2017", "Added DHS SIR button.", "Jodi Martin, Wright County")
call changelog_update ("11/16/2016", "Initial version.", "Jodi Martin, Wright County")



'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS------------------------------------------------------------------------------------------------------------------

'First Initial Dialog
BeginDialog Main_question_dlg, 0, 0, 226, 145, "Maintaining County"
  ButtonGroup ButtonPressed
    OkButton 75, 120, 50, 15
    CancelButton 135, 120, 50, 15
  EditBox 80, 30, 90, 15, prism_case_number
  Text 5, 60, 180, 20, "Are you requesting a county to maintain the case or are you responding to a request to maintain a case?"
  Text 15, 10, 105, 15, "Maintaining County CAAD note"
  DropListBox 35, 90, 115, 15, "Select one..."+chr(9)+"Requesting County"+chr(9)+"Responding County", script_run_mode
  Text 5, 35, 70, 10, "Prism case number"
  ButtonGroup ButtonPressed
    PushButton 145, 5, 60, 15, "DHS Sir-Milo Info", DHS_sir_button
EndDialog


'Requesting County Dialog
BeginDialog requesting_dlg, 0, 0, 256, 240, "Requesting County"
  EditBox 50, 45, 75, 15, Request_From
  EditBox 170, 45, 80, 15, Requesting_county
  EditBox 50, 65, 75, 15, Request_To
  EditBox 170, 65, 80, 15, Responding_county
  CheckBox 15, 90, 140, 10, "CP is now living in the reviewing county", CP_in_county
  CheckBox 15, 105, 140, 10, "CP is receiving PA in reviewing county", PA_reviewing
  CheckBox 15, 120, 175, 10, "Requesting County has not started any legal action", No_legal
  CheckBox 15, 135, 160, 10, "Reviewing County has an existing court order", Existing_order
  CheckBox 15, 150, 220, 10, "Reviewing County has the Companion case with controlling order", companion_case
  EditBox 90, 170, 145, 15, Additional_info
  EditBox 190, 190, 45, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 120, 215, 50, 15
    CancelButton 190, 215, 50, 15
  Text 5, 70, 45, 10, "Request To:"
  Text 30, 5, 175, 10, "Maintaining County Request and Review CAAD Note"
  Text 5, 175, 75, 10, "Additional information:"
  Text 140, 50, 25, 10, "County"
  Text 140, 195, 35, 10, "Signature:"
  Text 140, 70, 25, 10, "County"
  Text 5, 50, 50, 10, "Request From:  "
  Text 85, 25, 80, 10, "REQUESTING COUNTY"
EndDialog


BeginDialog responding_dlg, 0, 0, 236, 215, "Responding County"
  DropListBox 100, 30, 80, 15, "Select one..."+chr(9)+"Accepted"+chr(9)+"Denied", accept_deny
  EditBox 45, 55, 180, 15, Reason_note
  EditBox 135, 80, 75, 15, Transfer_to
  EditBox 75, 100, 40, 15, County_nbr
  EditBox 75, 120, 40, 15, Office_nbr
  EditBox 75, 140, 40, 15, Team_nbr
  EditBox 75, 160, 40, 15, Position_nbr
  EditBox 180, 175, 45, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 115, 195, 50, 15
    CancelButton 175, 195, 50, 15
  Text 15, 125, 55, 15, "Office number"
  Text 15, 145, 50, 15, "Team number"
  Text 15, 165, 55, 15, "Position number"
  Text 15, 60, 30, 10, "Reason:"
  Text 75, 10, 80, 15, "RESPONDING COUNTY"
  Text 135, 180, 35, 10, "Signature:"
  Text 40, 85, 80, 10, "Transfer to what county:"
  Text 15, 105, 50, 15, "County number"
  Text 25, 30, 75, 15, "Decision of request"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out and checks for PRISM case number
CALL check_for_PRISM(True)
CALL PRISM_case_number_finder(PRISM_case_number)

'The script will not run unless the mandatory fields are completed
DO
	Do
		DIALOG Main_question_dlg
		IF ButtonPressed = 0 then stopscript
		IF ButtonPressed = DHS_sir_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/PRISM/Documentation/Training/Job%20Aids/Forms/OnlyMaintainingCounty.aspx")
	Loop until buttonpressed = ok
	IF script_run_mode = "Select one..." THEN MsgBox "Please select a maintaining county action"
LOOP UNTIL script_run_mode <> "Select one..."

	IF script_run_mode = "Requesting County" THEN
	
'Goes to CAAD
	CALL Navigate_to_PRISM_screen ("CAAD")											'goes to the CAAD screen
PF5																	'F5 to add a note
EMWritescreen "A", 3, 29													'put the A on the action line

'Writes info from dialog into CAAD
EMWritescreen "T0098", 4, 54													'types T0098(CONTACT WITH WORKER FROM ANOTHER MN COUNTY)on caad code: line
EMWritescreen "Maintaining County", 16, 4								 	     		'types Maintaining County on the first line of the note
EMSetCursor 17, 4															' sets cursor on the 2nd line of the CAAD note

'Assures the mandatory fields are completed
DO
	err_msg = ""
	Dialog requesting_dlg
	IF ButtonPressed = 0 THEN StopScript		                                       		'Pressing Cancel stops the script
	CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF Request_From = "" THEN err_msg = err_msg & vbNewline & "You must enter the requesting county"
	IF Requesting_county = "" THEN err_msg = err_msg & vbNewline & "You must enter the requesting county"
  	IF Request_To = "" THEN err_msg = err_msg & vbNewline & "You must enter the requesting county"
	IF Responding_county = "" THEN err_msg = err_msg & vbNewline & "You must enter the responding county"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""												 	

'Writes info from dialog into CAAD
	CALL write_bullet_and_variable_in_CAAD("Requesting from CSO:", Request_From)
	CALL write_bullet_and_variable_in_CAAD("Requesting to CSO:", Request_To)
	CALL write_bullet_and_variable_in_CAAD("Requesting county:", Requesting_county)
	CALL write_bullet_and_variable_in_CAAD("Responding county:", Responding_county)
	If CP_in_county = checked then call write_variable_in_CAAD("CP is now living in the reviewing county")
	If PA_reviewing = 1 then call write_variable_in_CAAD("CP is receiving PA in reviewing county")
	If No_legal = 1 then call write_variable_in_CAAD("Requesting County has not started any legal action")
	If Existing_order = 1 then call write_variable_in_CAAD("Reviewing County has an existing court order")
	If companion_case = 1 then call write_variable_in_CAAD("Reviewing County has the companion case with controlling order")
	CALL write_bullet_and_variable_in_CAAD("Additional Info", Additional_info)
	CALL write_variable_in_CAAD(worker_signature)

	end if	
	
IF script_run_mode = "Responding County" THEN

'Assures the mandatory fields are completed
DO
	DIALOG responding_dlg
	IF ButtonPressed = stop_script_button THEN stopscript
	IF accept_deny = "Select one..." THEN MsgBox "Please select a maintaining county action"
LOOP UNTIL accept_deny <> "Select one..."

'Goes to CAAD
	CALL Navigate_to_PRISM_screen ("CAAD")											'goes to the CAAD screen
PF5																	'F5 to add a note
EMWritescreen "A", 3, 29													'put the A on the action line

'Writes info from dialog into CAAD
EMWritescreen "T0098", 4, 54													'types T0098(CONTACT WITH WORKER FROM ANOTHER MN COUNTY)on caad code: line
EMWritescreen "Maintaining County", 16, 4								 	     	 	'types Maintaining County on the first line of the note
EMSetCursor 17, 4															' sets cursor on the 2nd line of the CAAD note
	CALL write_bullet_and_variable_in_CAAD ("Maintaining request is", accept_deny)
	CALL write_bullet_and_variable_in_CAAD("Reason", Reason_note)
	CALL write_bullet_and_variable_in_CAAD ("Transfer to county", Transfer_to)
	CALL write_bullet_and_variable_in_CAAD ("County", County_nbr)
	CALL write_bullet_and_variable_in_CAAD ("Office", Office_nbr)
	CALL write_bullet_and_variable_in_CAAD ("Team", Team_nbr)
	CALL write_bullet_and_variable_in_CAAD ("Position", Position_nbr)
	IF Transfer_tocheck = 1 THEN CALL write_variable_in_CAAD("Transfer to county:", Transfer_to)
	CALL write_variable_in_CAAD(worker_signature)

	end if

script_end_procedure("")
