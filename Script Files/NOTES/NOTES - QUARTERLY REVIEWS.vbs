'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - QUARTERLY REVIEWS.vbs"
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

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
days_to_push_out_worklist = "90"	'This is the default

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog quarterly_reviews_dialog, 0, 0, 176, 85, "Quarterly Reviews Dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  EditBox 140, 25, 35, 15, days_to_push_out_worklist
  EditBox 70, 45, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 130, 10, "Days to push out worklist (default is 90):"
  Text 5, 50, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Sends a transmit to check for password issues
transmit

'Checking to make sure we're on USWT or USWD. If not the script will stop.
EMReadScreen worklist_check, 3, 21, 75
If worklist_check <> "USW" and worklist_check <> "CAW" then script_end_procedure("Worklist screen not found. Please start this script from the worklist you are trying to copy over.")

'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

Do
	Do
		Do
			dialog quarterly_reviews_dialog
			If buttonpressed = 0 then stopscript
			call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			If case_number_valid = False then MsgBox("Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''")
		Loop until case_number_valid = True
		If isnumeric(days_to_push_out_worklist) = False then MsgBox ("You must put a number in for the days to push out worklist.")
	Loop until isnumeric(days_to_push_out_worklist) = True


	EMReadScreen worklist_line_01, 72, 10, 4			'Reads worklist info, line by line
	EMReadScreen worklist_line_02, 72, 11, 4
	EMReadScreen worklist_line_03, 72, 12, 4
	EMReadScreen worklist_line_04, 72, 13, 4
	EMWriteScreen "__________", 17, 21				'clearing out worklist date
	EMWriteScreen days_to_push_out_worklist, 17, 52		'Adding the number of days to push out worklist
	EMWriteScreen "m", 3, 30					'Must modify the panel
	transmit
	call navigate_to_PRISM_screen("CAAD")
	pf5
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
Loop until case_activity_detail = "Case Activity Detail"

EMWriteScreen worklist_line_01, 16, 4	
EMWriteScreen worklist_line_02, 17, 4
EMWriteScreen worklist_line_03, 18, 4
EMWriteScreen worklist_line_04, 19, 4
EMWriteScreen "------" & worker_signature, 20, 4
EMWriteScreen "E0002", 4, 54

script_end_procedure("")
