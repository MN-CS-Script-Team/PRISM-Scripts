'Option Explicit -- COMMENTED OUT PER VKC REQUEST
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - PAYMENT PLAN REVIEW.vbs"
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


BeginDialog DLPP_dialog, 0, 0, 266, 200, "Payment Plan Dialog"
  EditBox 80, 10, 85, 15, PRISM_case_number
  DropListBox 80, 30, 95, 15, "Select one:"+chr(9)+"FIDM"+chr(9)+"Drivers License"+chr(9)+"Occupational License"+chr(9)+"Student Grant Hold", List1
  EditBox 70, 50, 55, 15, Beg_date
  EditBox 85, 70, 65, 15, total_due
  EditBox 75, 90, 70, 15, deliq_amt
  EditBox 135, 110, 125, 15, results
  CheckBox 5, 130, 75, 15, "Send DORD F0919", DLPP_checkbox
  CheckBox 5, 145, 145, 15, "Send Other Non-compliance DORD F0918", Other_checkbox
  EditBox 180, 155, 55, 15, workers_signature
  ButtonGroup ButtonPressed
    OkButton 145, 180, 50, 15
    CancelButton 200, 180, 50, 15
  Text 10, 15, 75, 10, "PRISM Case Number"
  Text 5, 35, 75, 15, "Type of Payment Plan"
  Text 5, 55, 65, 15, "Pay Plan Beg Date"
  Text 5, 75, 80, 15, "Payment Plan Total Due"
  Text 5, 95, 65, 15, "Delinquency Amount"
  Text 5, 115, 130, 10, "Results of Review and any action taken:"
  Text 105, 160, 70, 10, "Worker Signature:"
EndDialog

EMConnect ""


DO
		
Dialog DLPP_dialog
	IF ButtonPressed = 0 THEN StopScript
	IF workers_signature = "" THEN MsgBox "Please sign your CAAD note"
LOOP UNTIL workers_signature <> ""

	DO															'Making sure that user picks a type of plan other than select one
		IF LIST1 = "Select one" THEN MsgBox "Please select a Type of Payment Plan"			
	LOOP UNTIL LIST1 <> ""

DIM beta_agency

Dim DLPP_dialog, PRISM_case_number, List1, results, Beg_date, total_due, deliq_amt, DLPP_checkbox, Other_checkbox, workers_signature, ButtonPressed

Call check_for_PRISM(True)				'checks to make sure you are not passworded out

Call navigate_to_PRISM_screen("CAAD")		'goes to CAAD screen

PF5								'adding a caad note
EMWriteScreen "A", 3, 29							
EMWriteScreen "Free", 4, 54				'putting in caad code	
EMSetCursor 16, 4						'getting cursor in correct screen position to enter note.

'writes the CAAD note			
call write_variable_in_CAAD("Drivers License Payment Review")
call write_bullet_and_variable_in_CAAD("Type of Payment Plan: ", List1)
call write_bullet_and_variable_in_CAAD("Pay Plan Beg Date: ", Beg_date)
call write_bullet_and_variable_in_CAAD("Payment Plan Total Due: ", total_due)
call write_bullet_and_variable_in_CAAD("Delinquency Amount: ", deliq_amt)
call write_bullet_and_variable_in_CAAD("Results of Review and any actions taken: ", results)
IF DLPP_checkbox = 1 THEN
call write_variable_in_CAAD("Sent F0919 Non-Compliance")			'noting if this document was sent
END IF 
IF Other_checkbox = 1 THEN
call write_variable_in_CAAD("Sent OTHER F0918 Non-Compliance")		'noting if this document was sent
END IF
call write_variable_in_CAAD("*******")
call write_variable_in_CAAD(workers_signature)
transmit

IF DLPP_checkbox = 1 THEN 
call navigate_to_PRISM_screen("DORD")			'sending the document
PF5
EMWriteScreen "A", 3, 29
EmWriteScreen "F0919",  6, 36
transmit
END IF

IF Other_checkbox = 1 THEN 
call navigate_to_PRISM_screen("DORD")			'sending the document
PF5
EMWriteScreen "A", 3, 29					
EmWriteScreen "F0918", 6, 36
transmit
END IF


'Would like auto-filling for Beg Date, Total Due, Deliquency Amount 







