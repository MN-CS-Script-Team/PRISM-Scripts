'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "court-order-request.vbs"
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
CALL changelog_update("01/18/2017", "The worker signature field should now auto-populate.", "Kelly Hiestand, Wright County")
CALL changelog_update("11/30/2016", "The script has been updated to include a Requested Via drop down option of E-Filed. Signature Date has also been added to the order type field.", "Kelly Hiestand, Wright County")
CALL changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIMMING variables
DIM row, col, case_number_valid, Court_Order_Request_Dialog, prism_case_number, date_court_order_requested, requested_via_droplistbox, requested_from, court_order_number, create_worklist_checkbox, order_type, ButtonPressed


'THE DIALOG----------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog Court_Order_Request_Dialog, 0, 0, 406, 95, "Court Order Request"
  EditBox 80, 5, 70, 15, prism_case_number
  EditBox 290, 5, 65, 15, date_court_order_requested
  ComboBox 100, 30, 115, 15, "Click here to enter county name"+chr(9)+"CP"+chr(9)+"NCP", requested_from
  DropListBox 290, 30, 85, 15, "Select one..."+chr(9)+"E-Filed"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Inter-Office"+chr(9)+"Mail"+chr(9)+"SIR Email"+chr(9)+"Telephone", requested_via_droplistbox
  EditBox 80, 50, 85, 15, court_order_number
  EditBox 325, 50, 70, 15, order_type
  EditBox 80, 75, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 295, 75, 50, 15
    CancelButton 350, 75, 50, 15
  Text 235, 30, 50, 10, "Requested Via:"
  Text 5, 10, 70, 10, "Prism Case Number:"
  Text 5, 80, 70, 10, "Sign your CAAD note:"
  Text 190, 10, 95, 10, "Date Court Order Requested:"
  Text 235, 55, 90, 10, "Order Type/Signature Date:"
  Text 15, 55, 65, 10, "Court File Number:"
  Text 5, 30, 90, 20, "Requested From: (or type in County name)"
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

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

'The script will not run unless the CAAD note is signed, and you are in a valid Prism case, and makes requested from field mandatory
DO
	DO
		DO
			Dialog court_order_request_dialog
			IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
			IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                   'If worker sig is blank, message box pops saying you must sign caad note
			CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
		LOOP UNTIL worker_signature <> ""                                                            'Will keep popping up until worker signs note
		IF requested_from = "Select one..." THEN MsgBox "You must complete 'Requested From field'"   'Makes this field mandatory
	LOOP UNTIL requested_from <> "Select one..."
LOOP UNTIL case_number_valid = TRUE


'Navigates to CAAD and adds note
CALL navigate_to_PRISM_screen("CAAD")

'Adds a new CAAD note
PF5

EMWriteScreen "A", 3, 29


'Writes the CAAD NOTE
EMWriteScreen "B0170", 4, 54         'Type of Caad note
EMWriteScreen date_court_order_requested, 4, 37
EMSetCursor 16, 4
CALL write_bullet_and_variable_in_CAAD("Date Court Order Requested", date_court_order_requested)   'types date court order requested info
CALL write_bullet_and_variable_in_CAAD("Requested From", requested_from)                           'types requested from info
CALL write_bullet_and_variable_in_CAAD("Requested Via", requested_via_droplistbox)			   'types requested via info
CALL write_bullet_and_variable_in_CAAD("Court File Number", court_order_number)                    'types court file number info
CALL write_bullet_and_variable_in_CAAD("Order Type/Signature Date", order_type)                    'types order type info and signature date
CALL write_variable_in_CAAD(worker_signature)                                                      'types worker signature


'Saves the CAAD note
transmit

'Exits back out of that CAAD note
PF3

'Navigates to CAWT and creates worklist
CALL navigate_to_PRISM_screen("CAWT")

'Adds a new worklist
PF5

'Puts the A in the Action part
EMWriteScreen "A", 3, 30

'Writes the Worklist
EMWriteScreen "FREE", 4, 37
'Writes note in CAWT
EMWriteScreen "Did Court Order come in? See CAAD notes for request details.", 10, 4

script_end_procedure("")
