'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "waiver-of-personal-service.vbs"
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
CALL changelog_update("01/18/2017", "Worker Signature should now auto-populate.", "Kelly Hiestand, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIMMING variables
DIM row, col, case_number_valid, waiver_signed_date, prism_case_number, waiver_dialog, ButtonPressed

'THE DIALOG--------------------------------------------------------------------------------------------------

BeginDialog Waiver_Dialog, 0, 0, 236, 85, "Waiver of Personal Service"
  EditBox 80, 5, 75, 15, prism_case_number
  EditBox 180, 25, 55, 15, waiver_signed_date
  EditBox 80, 45, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 65, 50, 15
    CancelButton 180, 65, 50, 15
  Text 5, 10, 75, 10, "PRISM Case Number:"
  Text 5, 50, 70, 10, "Sign your CAAD Note:"
  Text 5, 30, 170, 15, "Date Waiver of Personal Service was signed by CP:"
EndDialog


'THE SCRIPT-------------------------------------------------------------------------------------------------

'Connects to Bluezone
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


'The script will not run unless the CAAD note is signed and there is a valid prism case number
DO
	DO
		Dialog waiver_dialog
		IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
		IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                   'If worker sig is blank, message box pops saying you must sign caad note
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
	LOOP UNTIL case_number_valid = True
LOOP UNTIL worker_signature <> ""                                                                  'Will keep popping up until worker signs note


'Navigates to CAAD and adds note
CALL navigate_to_PRISM_screen("CAAD")

'Adds a new CAAD note
PF5
EMWritescreen "A", 3, 29

'Writes the CAAD NOTE
EMWriteScreen "D5010", 4, 54     'Type of Caad note
EMSetCursor 16, 4
CALL write_bullet_and_variable_in_CAAD("Waiver of Personal Service Signed by CP", waiver_signed_date)
CALL write_variable_in_CAAD(worker_signature)
transmit  'Saves the CAAD note


script_end_procedure("")   'Stops the script
