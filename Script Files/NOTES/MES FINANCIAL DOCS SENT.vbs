option explicit

DIM beta_agency, row, col

'LOADING ROUTINE FUNCTIONS (FOR PRISM)--------------------------------------------------------------------------------------------------------------------------
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

'DIMMING VARIABLES---------------------------------------------------------------------------------------------------------------------------------------------------------

DIM financial_stmt_cp_check, financial_stmt_ncp_check, Cover_letter_cp_check, cover_letter_ncp_check, Waiver_cp_check, waiver_ncp_check, calendar_cp_check, calendar_ncp_check, past_support_cp_check, past_support_ncp_check, return_date, worker_signature, PRISM_case_number, MES_Financial_Docs_Sent_dialog, buttonpressed, case_number_valid

'Calling dialog for the MES Financial Docs Sent-----------------------------------------------------------------------------------------------------------------------------

BeginDialog MES_Financial_Docs_Sent_dialog, 0, 0, 271, 200, "MES Financial Docs Sent"
  EditBox 85, 5, 90, 15, prism_case_number
  CheckBox 20, 50, 75, 15, "Financial Statement", financial_stmt_cp_check
  CheckBox 150, 50, 75, 15, "Financial Statement", financial_stmt_ncp_check
  CheckBox 20, 65, 80, 15, "Cover Letter", Cover_letter_cp_check
  CheckBox 150, 65, 55, 15, "Cover Letter", cover_letter_ncp_check
  CheckBox 20, 80, 100, 15, "Waiver of Personal Service", Waiver_cp_check
  CheckBox 150, 80, 100, 15, "Waiver of Personal Service", waiver_ncp_check
  CheckBox 20, 95, 95, 15, "Parenting Time Calendar", calendar_cp_check
  CheckBox 150, 95, 95, 15, "Parenting Time Calendar", calendar_ncp_check
  CheckBox 20, 110, 70, 15, "Past Support Form", Past_support_cp_check
  CheckBox 150, 110, 90, 15, "Past Support Form", past_support_cp_check
  EditBox 140, 135, 60, 15, return_date
  EditBox 140, 155, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 180, 50, 15
    CancelButton 195, 180, 50, 15
  GroupBox 140, 35, 115, 95, "Documents Sent to NCP:"
  Text 65, 160, 70, 10, "Sign your CAAD Note:"
  Text 55, 140, 80, 10, "Requested Return Date:"
  GroupBox 10, 35, 115, 95, "Documents Sent to CP:"
  Text 15, 10, 70, 10, "Prism Case Number:"
EndDialog


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
		dialog MES_Financial_Docs_Sent_dialog
		IF buttonpressed = 0 THEN stopscript
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
		IF worker_signature = "" THEN MSGbox "Please sign your CAAD Note"						'if the signature is blank pop up a message box
		IF return_date = "" THEN MSGbox "Please enter a Requested Return Date"				      'if the date field is blank pop up a message box
		IF IsDate(return_date) = False THEN MsgBox "You must enter a valid date"	      	      'makes sure the date field is a valid date
	LOOP UNTIL case_number_valid = True
LOOP UNTIL worker_signature <> "" and return_date <> "" and IsDate(return_date) = TRUE                      'tells the loop to keep running until the date and signature fields are filled in and the date is valid.  (if you have a Do stmt, you must have a LOOP UNTIL stmt)


'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

				
'Goes to CAAD
CALL Navigate_to_PRISM_screen ("CAAD")										'goes to the CAAD screen
PF5																'F5 to add a note
EMWritescreen "A", 3, 29												'put the A on the action line

'Writes info from dialog into CAAD
EMWritescreen "FREE", 4, 54											      	'types free on caad code type
EMWritescreen "*MES Financial Docs Sent*", 16, 4									'types title of the free caad on the first line of the note
EMWriteScreen "Docs Sent to CP:", 17, 4
EMSetCursor 18, 4
IF financial_stmt_cp_check = 1 THEN CALL write_variable_in_CAAD("Financial Statement")  	       'putting the info that is checked from the dialog box into the caad if it is checked
IF cover_letter_cp_check = 1 THEN CALL write_variable_in_CAAD("Cover Letter")
IF waiver_cp_check = 1 THEN CALL write_variable_in_CAAD("Waiver of Personal Service")
IF calendar_cp_check = 1 THEN CALL write_variable_in_CAAD("Parenting Time Calendar")
If past_support_cp_check = 1 THEN CALL write_variable_in_CAAD("Past Support Form")


'EMWriteScreen "Docs Sent to NCP:"  'NEED TO FIGURE THIS OUT- I cannot get this to write in the CAAD note before any of these NCP ones are checked.
IF financial_stmt_ncp_check = 1 THEN CALL write_variable_in_CAAD("Financial Statement")     	'putting the info that is checked from the dialog box into the caad if it is checked
IF cover_letter_ncp_check = 1 THEN CALL write_variable_in_CAAD("Cover Letter")
IF waiver_ncp_check = 1 THEN CALL write_variable_in_CAAD("Waiver of Personal Service")
IF calendar_ncp_check = 1 THEN CALL write_variable_in_CAAD("Parenting Time Calendar")
If past_support_ncp_check = 1 THEN CALL write_variable_in_CAAD("Past Support Form")


CALL write_bullet_and_variable_in_CAAD("Requested return date", return_date)
CALL write_variable_in_CAAD(worker_signature)
transmit
PF3

'Goes to CAWT to add a FREE worklist for the CP's FIN STMT DUE, HELP!  this is getting weird in CAWT if the CP financial stmt variable is not checked.
IF financial_stmt_cp_check = 1 THEN CALL navigate_to_PRISM_screen ("CAWT")
PF5
EMWriteScreen "A", 3, 30
EMWriteScreen "FREE", 4, 37												'types free on worklist item: line
EMWriteScreen "CP's Financial Stmt Due", 10, 4 						      'types description, have docs been returned   
EMWriteScreen return_date, 17, 21      			                  					'types 30 in the calendar days field
transmit
PF3

'Goes to CAWT to add a FREE worklist for the NCP's FIN STMT DUE   *NEED TO FIX!  It is automatically creating this worklist even if the NCP Financial Stmt variable is not checked on the dialog box, I only want this worklist to be created if that specific box is checked.
IF financial_stmt_ncp_check = 1 THEN CALL navigate_to_PRISM_screen ("CAWT")
PF5
EMWriteScreen "A", 3, 30
EMWritescreen "FREE", 4, 37												'types free on worklist item: line
EMWritescreen "NCP's Financial Stmt Due", 10, 4 						            'types description, have docs been returned   
EMWritescreen return_date, 17, 21      			                  			      'types 30 in the calendar days field
transmit
PF3


script_end_procedure("")                                                                     	'stopping the script






