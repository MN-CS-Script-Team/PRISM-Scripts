'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "csenet-info.vbs"
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
CALL changelog_update("01/18/2017", "The worker signature should now auto-populate.", "Kelly Hiestand, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIMMING VARIABLES-------------------------------------------------------------------------------------------------------------------------------------
DIM prism_case_number, csenet_total, csenet_info_dialog, ButtonPressed, write_new_line_in_CAAD, csenet_sent_recd, reason_code_line, row, col, beta_agency, case_number_valid, csenet_dateline, csenet_line_01, csenet_line_02, csenet_line_03, csenet_line_04, csenet_line_05

'THE DIALOG-------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog csenet_info_dialog, 0, 0, 216, 90, "CSENET Info"
  EditBox 85, 20, 60, 15, prism_case_number
  EditBox 85, 45, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 105, 70, 50, 15
    CancelButton 160, 70, 50, 15
  Text 10, 50, 70, 10, "Sign your CAAD note:"
  Text 10, 25, 70, 15, "Prism Case Number:"
  Text 10, 5, 235, 10, "Make sure INTD message is open before running this script."
EndDialog


'THE SCRIPT-----------------------------------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Brings Bluezone to the Front
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
		Dialog csenet_info_dialog
		IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
		IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                   'If worker sig is blank, message box pops saying you must sign caad note
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
	LOOP UNTIL case_number_valid = True
LOOP UNTIL worker_signature <> ""                                                                  'Will keep popping up until worker signs note


'Reads the contents of the CSENET for CAAD noting
EMReadScreen reason_code_line, 45, 13, 14
EMReadScreen csenet_sent_recd, 1, 14, 61
EMReadScreen csenet_line_01, 80, 15, 2
EMReadScreen csenet_line_02, 80, 16, 2
EMReadScreen csenet_line_03, 80, 17, 2
EMReadScreen csenet_line_04, 80, 18, 2
EMReadScreen csenet_line_05, 80, 19, 2

csenet_line_01 = replace(csenet_line_01, "_", "")
csenet_line_02 = replace(csenet_line_02, "_", "")
csenet_line_03 = replace(csenet_line_03, "_", "")
csenet_line_04 = replace(csenet_line_04, "_", "")
csenet_line_05 = replace(csenet_line_05, "_", "")

csenet_total = csenet_line_01 & " " & csenet_line_02 & " " & csenet_line_03 & " " & csenet_line_04 & " " & csenet_line_05

'Navigates to CAAD and adds the note
CALL navigate_to_PRISM_screen("CAAD")


'Adds new CAAD note
PF5

EMSetCursor 16, 4
CALL write_variable_in_CAAD("##CSENET INFO##")                                      'Writes CSENET INFO in title

CALL write_bullet_and_variable_in_CAAD("CSESNET sent/rcd", csenet_sent_recd)         'Writes CSENET Sent/Recd and Date/Time
CALL write_bullet_and_variable_in_CAAD("Reason Code", reason_code_line)             'Writes in the reason code
CALL write_bullet_and_variable_in_CAAD("CSENET Comments", csenet_total)             'Writes CSENET Comments
CALL write_variable_in_CAAD("---")				                              'Writes Worker Signature
CALL write_variable_in_CAAD(worker_signature)


EMWriteScreen "A", 3, 29                                                          'Writes A to add the new caad note

'Writes the CAAD note type
EMWriteScreen "T0111", 4, 54

EMWriteScreen PRISM_case_number, 4, 8

'Saves the CAAD note
transmit

'Exits back out of that CAAD note
PF3

script_end_procedure("")             'Stops the script
