'Gathering stats==========================================
name_of_script = "NOTES - CASE TRANSFER.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'DIMMING VARIABLES---------------------------------------------------------------------------------------------------------------------------------------------


'DIALOG BOX----------------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog Case_Transfer_dialog, 0, 0, 316, 165, "Case Transfer"
  EditBox 85, 5, 80, 15, prism_case_number
  EditBox 50, 45, 35, 15, county
  EditBox 50, 65, 35, 15, office
  EditBox 50, 85, 35, 15, Team:
  EditBox 50, 105, 35, 15, Position
  EditBox 130, 65, 175, 15, transfer_reason
  CheckBox 130, 90, 115, 15, "Sent New Worker Letter to CP", letter_checkbox
  EditBox 205, 115, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 140, 50, 15
    CancelButton 245, 140, 50, 15
  Text 20, 90, 25, 10, "Team:"
  Text 130, 120, 75, 10, "Sign your CAAD note:"
  Text 10, 10, 70, 10, "PRISM Case Number:"
  Text 20, 70, 20, 10, "Office:"
  Text 15, 110, 30, 10, "Position:"
  Text 130, 50, 60, 10, "Transfer Reason:"
  Text 20, 50, 25, 10, "County:"
  GroupBox 5, 30, 105, 105, "Transfer To:"
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------

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

'The script will not run unless the CAAD note is signed, and you are in a valid Prism case

office = "001"

DO
	DO
		DO
			Dialog Case_Transfer_dialog
			IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
			IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                   'If worker sig is blank, message box pops saying you must sign caad note
			CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
			IF transfer_reason = "" THEN MsgBox "You must type a Transfer Reason!"                 'It will loop until transfer reason is filled out.
		LOOP UNTIL transfer_reason <> ""	
	LOOP UNTIL worker_signature <> ""                             'Will keep popping up until worker signs note
LOOP UNTIL case_number_valid = TRUE

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
CALL write_variable_in_CAAD("**Case Transfer**")
CALL write_bullet_and_variable_in_CAAD("Transferred To", position_name & " at " & office_name) 
CALL write_bullet_and_variable_in_CAAD("Transfer Reason", Transfer_reason)
IF letter_checkbox = 1 THEN CALL write_variable_in_CAAD("* Sent New Worker letter to CP")
CALL write_variable_in_CAAD(worker_signature) 
transmit

script_end_procedure("Case has been transferred and CAAD noted. If you use OnBase as your EDMS System, don't forget to SCREEN SCRAPE if necessary.")
