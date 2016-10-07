'GATHERING STATS----------------------------------------------------------------------------------------------------

name_of_script = "ACTIONS - Unreimbursed Uninsured Docs.vbs"
start_time = timer


'this is a function document
DIM beta_agency 'remember to add

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
'this is where the copy and paste from functions library ended

'DIALOGS---------------------------------------------------------------------------
DIM UnUn_Dialog, PRISM_case_number, CP, NCP, Percent, err_msg, ButtonPressed, case_number_is_valid

BeginDialog UnUn_Dialog, 0, 0, 291, 145, "Unreimbursed Uninsured Docs"
  EditBox 60, 45, 90, 15, PRISM_case_number
  CheckBox 50, 85, 20, 10, "CP", CP
  CheckBox 120, 85, 25, 10, "NCP", NCP
  EditBox 175, 100, 25, 15, Percent
  ButtonGroup ButtonPressed
    OkButton 180, 125, 50, 15
    CancelButton 235, 125, 50, 15
  Text 25, 10, 240, 15, "This script will gernerate DORD DOCS F0944, F0659, and F0945 for collection of Unreimbursed and Uninsured Medical and Dental Expenses."
  Text 5, 50, 50, 10, "Case Number"
  Text 5, 70, 175, 10, "Check who requested Unreimbursed/Uninsured forms"
  Text 90, 85, 15, 10, "or"
  Text 5, 105, 165, 10, "Enter the PERCENT owed by non requesting party:"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'brings me to the CAPS screen
CALL navigate_to_PRISM_screen ("CAPS")

'this auto fills prism case number in dialog
EMReadScreen PRISM_case_number, 13, 4, 8 

'THE LOOP--------------------------------------
'adding a loop
Do
	err_msg = ""
	Dialog UnUn_Dialog 'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed"
		IF Percent = "" THEN err_msg = err_msg & vbNewline & "Percent of Unreimbursed Uninsured Expense must be completed."
		'IF both cp box and ncp box blank
		IF CP = 0 AND NCP = 0 THEN err_msg = vbNewline & "Either CP or NCP must be selected."
		IF CP = 1 AND NCP = 1 THEN err_msg = vbNewline & "Either CP or NCP must be selected."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'END LOOP--------------------------------------


'creates DORD doc for NCP
IF NCP = checked THEN

	CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0945", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0944", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0659", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	'shift f2, to get to user lables
	PF14
	EMWriteScreen "u", 20,14
	transmit
	EMSetCursor 7, 5
	EMWriteScreen "S", 7, 5

	transmit
	EMWriteScreen (Percent), 16, 15
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


EMWriteScreen "B", 3, 29
transmit

END IF

'creates DORD doc for CP
IF CP = checked THEN

	CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0945", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0944", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0659", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	'shift f2, to get to user labels
	PF14
	EMWriteScreen "u", 20,14
	transmit
	EMSetCursor 7, 5
	EMWriteScreen "S", 7, 5

	'enters the percent typed in the dialog box
	transmit
	EMWriteScreen (Percent), 16, 15
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit

'''need to select legal heading
DIM LH_dialog

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


EMWriteScreen "B", 3, 29
transmit

END IF


script_end_procedure("")


