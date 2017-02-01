'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "financial-statement-follow-up.vbs"
start_time = timer


'LOADING ROUTINE FUNCTIONS (FOR PRISM)--- UPDATED 9/8/16 to MASTER FUNCLIB--------------------------------------------------------------
IF IsEmpty(FuncLib_URL) = TRUE THEN 'Shouldn't load FuncLib if it already loaded once
    IF run_locally = FALSE or run_locally = "" THEN    'If the scripts are set to run locally, it skips this and uses an FSO below.
        IF use_master_branch = TRUE THEN               'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        Else                                            'Everyone else should use the release branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        End if
        SET req = CreateObject("Msxml2.XMLHttp.6.0")                'Creates an object to get a FuncLib_URL
        req.open "GET", FuncLib_URL, FALSE                          'Attempts to open the FuncLib_URL
        req.send                                                    'Sends request
        IF req.Status = 200 THEN                                    '200 means great success
            Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
            Execute req.responseText                                'Executes the script code
        ELSE                                                        'Error message
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
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")
				
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


BeginDialog finacial_statement_dialog, 0, 0, 231, 100, "Finacial Statement Follow up"
  Text 10, 10, 70, 10, "PRISM Case Number"
  EditBox 85, 5, 130, 15, PRISM_case_number
  Text 10, 30, 75, 10, "Select your recipient:"
  DropListBox 90, 30, 110, 20, "CPP - Custodial Parent"+chr(9)+"NCP - Noncustodial Parent"+chr(9)+"BOTH - CP and NCP", recipient_code
  Text 10, 55, 70, 10, "Worker's Signuature"
  EditBox 85, 50, 130, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 130, 75, 40, 15
    CancelButton 175, 75, 40, 15
EndDialog

'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

call PRISM_case_number_finder(PRISM_case_number)

Do
	err_msg = ""
	Dialog finacial_statement_dialog
	cancel_confirmation
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	IF worker_signature = "" THEN err_msg = err_msg & vbNEWline & "You must sign your CAAD note"
	IF recipient_code = "" THEN err_msg = err_msg & vbNEWline & "You must select a Recipient!"
	IF err_msg <> "" THEN MsgBox "***Notice***" & vbNEWline & err_msg &vbNEWline & vbNEWline & "Please resolve for the script"
LOOP UNTIL err_msg = ""
	transmit
	EMReadScreen PRISM_check, 5, 1, 36

Do
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


'************FUNCTIONS **************************
IF recipient_code = "CPP - Custodial Parent" or recipient_code = "NCP - Noncustodial Parent" THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0104", 6, 36
	EMWriteScreen Left(recipient_code, 3), 11, 51
	EMWriteScreen "        ", 4, 50
	EMWriteScreen "       ", 4, 59
	transmit

	PF14
	
	EMWriteScreen "U", 20, 14
	transmit

	dord_row = 7
	DO
		EMWriteScreen "S", dord_row, 5
		dord_row = dord_row + 1
	LOOP UNTIL dord_row = 18
	transmit

	EMWriteScreen "Within the last 30 days, our office forwarded you a packet", 16, 15
	transmit
	EMWriteScreen "of information to assist us in either establishing a child", 16, 15
	transmit
	EMWriteScreen "support obligation or reviewing a current support order", 16, 15
	transmit
	EMWriteScreen "for modification action.  You were asked to complete and ", 16, 15
	transmit
	EMWriteScreen "return the paperwork to our office.  As of the date of ", 16, 15
	transmit
	EMWriteScreen "this letter, we have not received said information from ", 16, 15
	transmit
	EMWriteScreen "you. If we do not receive the completed paperwork back ", 16, 15
	transmit
	EMWriteScreen "from you on or before noon, "& (DateAdd("d",7,Date)) &" we will", 16, 15
	transmit
	EMWriteScreen " move forward and be forced to utilize our own resources", 16, 15
	transmit
	EMWriteScreen "to determine your current financial situation. Therefore,", 16, 15
	transmit
	EMWriteScreen "we do await your prompt response.", 16, 15
	transmit
	transmit

	PF3 'goes back to main dord screen

	EMWriteScreen "M", 3, 29 'modify the document
	transmit

	PF9	'print the document
	transmit
	PF3

CALL navigate_to_PRISM_screen("CAAD")
EMWriteScreen "D", 8, 005
transmit
EMWriteScreen "M", 3, 029


EMSetCursor 16, 004

Call Write_variable_in_CAAD("Finacial Statement follow up letter send to " & Left(recipient_code, 3))
transmit


END IF

IF recipient_code = "BOTH - CP and NCP" THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0104", 6, 36
	EMWriteScreen "CPP", 11, 51
	EMWriteScreen "        ", 4, 50
	EMWriteScreen "       ", 4, 59
	transmit

	PF14
	
	EMWriteScreen "U", 20, 14
	transmit

	dord_row = 7
	DO
		EMWriteScreen "S", dord_row, 5
		dord_row = dord_row + 1
	LOOP UNTIL dord_row = 18
	transmit

	EMWriteScreen "Within the last 30 days, our office forwarded you a packet", 16, 15
	transmit
	EMWriteScreen "of information to assist us in either establishing a child", 16, 15
	transmit
	EMWriteScreen "support obligation or reviewing a current support order", 16, 15
	transmit
	EMWriteScreen "for modification action.  You were asked to complete and ", 16, 15
	transmit
	EMWriteScreen "return the paperwork to our office.  As of the date of ", 16, 15
	transmit
	EMWriteScreen "this letter, we have not received said information from ", 16, 15
	transmit
	EMWriteScreen "you. If we do not receive the completed paperwork back ", 16, 15
	transmit
	EMWriteScreen "from you on or before noon, "& (DateAdd("d",7,Date)) &" we will", 16, 15
	transmit
	EMWriteScreen " move forward and be forced to utilize our own resources", 16, 15
	transmit
	EMWriteScreen "to determine your current financial situation. Therefore,", 16, 15
	transmit
	EMWriteScreen "we do await your prompt response.", 16, 15
	transmit
	transmit

	PF3 'goes back to main dord screen

	EMWriteScreen "M", 3, 29 					'modify the document
	transmit

	PF15						'creating duplicate document
	EMWriteScreen Left(PRISM_case_number, 10), 10, 020	'adding the case number
	EMWriteScreen Right(PRISM_case_number, 2), 10, 031
	Transmit								'Back to DORD Screen
	PF3									'Back to DORD Screen
	EMWriteScreen "M", 3, 29					'Modifying the document
	EMWriteScreen "NCP", 11, 51					'changing to NCP document
	Transmit
	PF9									'printing
	Transmit
	EMWriteScreen "B", 3, 29
	Transmit							'Browsing for CP Letter
	EMWriteScreen "S", 5, 5						'Selecting CP Letter
	Transmit
	PF9									'Printing CP Letter
	Transmit

CALL navigate_to_PRISM_screen("CAAD")							'modifying and noting the caad note that the letter was sent
EMWriteScreen "D", 8, 005
transmit
EMWriteScreen "M", 3, 029
EMSetCursor 16, 004
Call Write_variable_in_CAAD("Finacial Statement follow up letter send to CP")
transmit
PF3
EMWriteScreen "D", 9, 005
Transmit
EMWriteScreen "M", 3, 029
EMSetCursor 16, 004
Call Write_variable_in_CAAD("Finacial Statement follow up letter send to NCP")
Transmit
PF3

END IF


script_end_procedure("")

script_end_procedure("")
