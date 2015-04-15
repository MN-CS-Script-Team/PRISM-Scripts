option explicit

'LOADING ROUTINE FUNCTIONS (FOR PRISM)-------------------------------------------------------------------------------------------------------------------
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
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF


'DIMMING variables
DIM beta_agency, row, col, worker_signature, ButtonPressed, Fee_Suppression_dialog, PRISM_case_number, case_number_valid, case_number, one_year_from_now

'THE DIALOG----------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog Fee_Suppression_dialog, 0, 0, 156, 115, "Fee Code Suppression"
  EditBox 45, 40, 75, 15, PRISM_Case_number
  EditBox 65, 60, 55, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 90, 50, 15
    CancelButton 100, 90, 50, 15
  Text 15, 45, 30, 10, "Case #:"
  Text 25, 65, 35, 10, "Signature:"
  Text 5, 5, 140, 25, "This will manually override the Fee Code Suppression setting the date for 1 year later than today."
EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to Bluezone
EMConnect ""			

'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

'Makes sure worker is in a valid PRISM Case, and workers signs caad note.
DO
	DO
		dialog Fee_Suppression_dialog
		IF buttonpressed = 0 THEN stopscript
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
		IF worker_signature = "" THEN MSGbox "Please sign your CAAD Note"				'if the signature is blank pop up a message box
	LOOP UNTIL case_number_valid = True
LOOP UNTIL worker_signature <> "" 							                     	'tells the loop to keep running until the signature field is filled in


'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

'Goes to CAST screen and PF11 over 							
CALL Navigate_to_PRISM_screen ("CAST")										'navigates to CAST
PF11																'Presses PF1 to move right 1 screen
EMWritescreen "M", 3,29													'puts M on the action line

'Updates State Fee Cd: to M in order to suppress the 2% fee
EMWritescreen "M", 9, 17												'changes State Fee Cd: code to M
EMSetCursor 10, 17													'puts cursor on fee elig date line									



'*TROUBLE RIGHT HERE** I'm trying to change the date on this field (the Fee Elig Date: field) to be 1 year from today's date, but I keep getting errors
'this script is harder to test in the training region as the fee code's on the CAST screen are not editable.  I have been testing it on a case that is closing in production region
DateAdd "d", 365, date													'changes the Fee eligible date to 1 year from today's date



transmit

'Writes info into CAAD
EMWritescreen "FREE", 4, 54	
EMWritescreen "Cost Recovery Fee Override", 16, 4								'types title of the free caad on the first line of the note
EMSetCursor 17, 4														'puts the cursor on the very next line to be ready to enter the not
EMWritescreen "Supervisor overrode cost recovery fee.  Case is NPA due to MNSure interface issue."
call write_variable_in_CAAD(worker_signature)									'adds worker initials from dialog box
transmit
PF3

script_end_procedure("")                                                                     	'stopping the script












