'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "fee-suppression-override.vbs"
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
Call changelog_update("01/20/2017", "Worker signature should now auto-populate.", "Kelly Hiestand, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIMMING variables
DIM row, col, ButtonPressed, Fee_Suppression_dialog, PRISM_case_number, Fee_date, CAAD_standard_checkbox,CAAD_text_checkbox, CAAD_text

'THE DIALOG----------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog Fee_Suppression_dialog, 0, 0, 346, 170, "Fee Code Suppression"
  EditBox 45, 25, 75, 15, PRISM_Case_number
  EditBox 65, 45, 50, 15, Fee_date
  CheckBox 15, 90, 320, 10, "Suppressed cost recovery fee.  Case is NPA due to ongoing METS interface issue.  ", CAAD_standard_checkbox
  CheckBox 15, 110, 90, 10, "Enter text for CAAD note.", CAAD_text_checkbox
  EditBox 110, 105, 225, 15, CAAD_text
  EditBox 75, 130, 55, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 235, 150, 50, 15
    CancelButton 290, 150, 50, 15
  Text 5, 10, 280, 10, "This script will manually override the Fee Code Suppression and create a CAAD note."
  Text 15, 30, 30, 10, "Case #:"
  Text 15, 50, 50, 10, "Fee Elig Date: "
  Text 125, 50, 70, 10, "(format 01/01/2001)"
  Text 5, 75, 105, 10, "Select CAAD note option below."
  Text 15, 135, 60, 10, "Worker Signature"
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to Bluezone
EMConnect ""

'brings me to the CAPS screen
CALL navigate_to_PRISM_screen ("CAST")

'this auto fills prism case number in dialog
EMReadScreen PRISM_case_number, 13, 4, 8


'adding LOOP to make sure info in dialog box is entered correctly
DO
	err_msg = ""
	dialog Fee_Suppression_dialog
	IF buttonpressed = 0 THEN stopscript
	IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "Please sign your CAAD Note"				'if the signature is blank pop up a message box
	IF Fee_date = "" THEN err_msg = err_msg & vbNewline & "Fee Eligibility end date must be entered."
	IF CAAD_standard_checkbox = 0 AND CAAD_text_checkbox = 0 THEN err_msg = err_msg & vbNewline & "Please select one CAAD note option."
	IF CAAD_standard_checkbox = 1 AND CAAD_text_checkbox = 1 THEN err_msg = err_msg & vbNewline & "Please select only one CAAD note option."
	IF CAAD_text_checkbox = 1 AND CAAD_text = "" THEN err_msg = err_msg & vbNewline & "Please enter the text for your CAAD note."
 	IF err_msg <> "" THEN
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF

LOOP UNTIL err_msg = ""

'END LOOP

'to pull up my prism
EMFocus

'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)


'Goes to CAST screen and PF11 over
CALL navigate_to_PRISM_screen("CAST")
PF11

'fixes date to the correct format xx/xx/xxxx
CALL create_mainframe_friendly_date(FEE_date, 10, 17, "YYYY")


'Updates State Fee Cd: to M in order to suppress the 2% fee and adds date
EMWritescreen "M", 9, 17
EMSetCursor 10, 17
EMWritescreen FEE_date, 10, 17
EMWritescreen "M", 3, 29
transmit

EMReadScreen mod_success,21 , 24, 22
IF mod_success <> "modified successfully" THEN
	Msgbox "Cast was not modified correctly, please reneter information correctly.  Script Ended."
	StopScript
END IF

'Writes info into CAAD for standard note
IF CAAD_standard_checkbox = 1 THEN
	CALL Navigate_to_PRISM_screen ("CAAD")										'navigates to CAADescreen "FREE", 4, 54												'types title of the free caad on the first line of the note
	PF5
	EMWriteScreen "Free", 4, 54
	EMWriteScreen "*Cost Recovery Fee Override*", 16, 4								'writes this as a title line for the caad note.
	EMSetCursor 17, 4
	CALL write_variable_in_CAAD ("Per state recommendation, suppressed cost recovery fee until "  &  Fee_date &  ".  Case is NPA due to ongoing METS interface issue.")
	CALL write_variable_in_CAAD(worker_signature)							  		'adds worker initials from dialog box
	transmit
	PF3
END IF

IF CAAD_text_checkbox = 1 THEN
	CALL Navigate_to_PRISM_screen ("CAAD")										'navigates to CAADescreen "FREE", 4, 54												'types title of the free caad on the first line of the note
	PF5
	EMWriteScreen "Free", 4, 54
	EMWriteScreen "*Cost Recovery Fee Override*", 16, 4								'writes this as a title line for the caad note.
	EMSetCursor 17, 4
	CALL write_variable_in_CAAD(CAAD_text)
	CALL write_variable_in_CAAD(worker_signature)							  		'adds worker initials from dialog box
	transmit
	PF3
END IF

'Goes to CAST screen and PF11 over
CALL navigate_to_PRISM_screen("CAST")
PF11

script_end_procedure("")                                                                     	'stopping the script
