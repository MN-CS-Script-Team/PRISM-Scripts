'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - WAIVER OF PERSONAL SERVICE.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED

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


'DIMMING variables
DIM beta_agency, row, col, case_number_valid, waiver_signed_date, prism_case_number, worker_signature, waiver_dialog, ButtonPressed

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





