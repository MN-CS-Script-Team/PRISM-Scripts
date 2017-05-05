'Gathering stats-------------------------------------------------------------------------------------
name_of_script = "case-transfer.vbs"
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
call changelog_update("05/02/2017", "Added additional information to the last message box.", "Jodi Martin, Wright County")
CALL changelog_update("01/18/2017", "A bug was fixed in this script to require the team field have 3 characters.", "Kelly Hiestand, Wright County")
call changelog_update("11/16/2016", "County, Office, Team and Position fields now have length requirements.", "Kelly Hiestand, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")




'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE DIALOG------------------------------------------------------------------------------------------
'Single case transfer dialog is here
BeginDialog Case_Transfer_dialog, 0, 0, 316, 160, "Case Transfer"
  EditBox 85, 5, 80, 15, prism_case_number
  EditBox 50, 45, 35, 15, county
  EditBox 50, 65, 35, 15, office
	EditBox 50, 85, 35, 15, team
  EditBox 50, 105, 35, 15, Position
  EditBox 130, 45, 175, 15, transfer_reason
  DropListBox 190, 70, 85, 15, "Select One..."+chr(9)+"Internal"+chr(9)+"External", transfer_type
  CheckBox 130, 95, 115, 15, "Sent New Worker Letter to CP", letter_checkbox
  EditBox 205, 115, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 140, 50, 15
    CancelButton 255, 140, 50, 15
  Text 130, 120, 75, 10, "Sign your CAAD note:"
  Text 10, 10, 70, 10, "PRISM Case Number:"
  Text 20, 70, 20, 10, "Office:"
  Text 15, 110, 30, 10, "Position:"
  Text 130, 35, 60, 10, "Transfer Reason:"
  Text 20, 50, 25, 10, "County:"
  GroupBox 5, 30, 105, 105, "Transfer To:"
  Text 130, 75, 55, 10, "Type of Transfer:"
  Text 20, 90, 25, 10, "Team:"
EndDialog

'THE SCRIPT---------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Brings Bluezone to the front
EMFocus

'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

'Makes sure you are not passworded out
CALL check_for_PRISM(True)
DO
	err_msg = ""
	Dialog Case_Transfer_dialog
	IF ButtonPressed = 0 THEN StopScript
	CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = False THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF len(county)<> 3 then err_msg = err_msg & vbNewline & "You must enter a 3 character county code!"
	If len(office)<> 3 then err_msg = err_msg & vbNewline & "You must enter a 3 character office code!"
	If len(team)<> 3 then err_msg = err_msg & vbNewline & "You must enter a 3 character team code!"
	If len(position)<> 2 then err_msg = err_msg & vbNewline & "You must enter a 2 character position code!"
	IF transfer_reason = "" THEN err_msg = err_msg & VbNewline & "You must type a Transfer Reason!"
	IF transfer_type = "Select One..." THEN err_msg = err_msg & VbNewline & "You must select the Type of Transfer!"
	IF worker_signature = "" THEN err_msg = err_msg & VbNewline & "You must sign your CAAD note!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'Navigates to CAAS screen to transfer case
CALL navigate_to_PRISM_screen("CAAS")

'Writes an M on CAAS to modify info on screen
EMWriteScreen "M", 3, 29

EMWriteScreen county, 9, 20
EMWriteScreen office, 10, 20
EMWriteScreen team, 11, 20
EMWriteScreen position, 12, 20

transmit

EMReadScreen office_name, 34, 10, 25
EMReadScreen position_name, 20, 12, 25

position_name = trim(position_name)
office_name = trim(office_name)


'Navigates to CAAD and adds note
CALL navigate_to_PRISM_screen("CAAD")

'Adds a new CAAD note
PF5

EMWriteScreen "A", 3, 29


'Writes the CAAD NOTE
EMWriteScreen "FREE", 4, 54      'Types FREE on type of CAAD line
EMSetCursor 16, 4
CALL write_variable_in_CAAD(" ** " & transfer_type & " Case Transfer **")
CALL write_bullet_and_variable_in_CAAD("Transferred To", position_name & " at " & office_name)
CALL write_bullet_and_variable_in_CAAD("Transfer Reason", Transfer_reason)
IF letter_checkbox = 1 THEN CALL write_variable_in_CAAD("* Sent New Worker letter to CP")
CALL write_variable_in_CAAD(worker_signature)
transmit

MsgBox "Success! The case has been transferred and CAAD noted. Please complete any additional necessary actions per your county's case transfer process."

script_end_procedure("")
