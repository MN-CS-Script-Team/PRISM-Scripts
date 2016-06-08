'GATHERING STATS----------------------------------------------------------------------------------------------------

name_of_script = "ACTIONS - Unreimbursed Uninsured Docs.vbs"
start_time = timer


'this is a function document
DIM beta_agency 'remember to add

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

'this is where the copy and paste from functions library ended


'DIALOGS---------------------------------------------------------------------------
DIM UnUn_Dialog, PRISM_case_number, CP, NCP, Percent, err_msg, ButtonPressed, case_number_is_valid, amount, important_checkbox, CAAD_checkbox, Enforce_checkbox, Aff_Service_checkbox, worker_signature

BeginDialog UnUn_Dialog, 0, 0, 296, 310, "Unreimbursed Uninsured Docs"
  EditBox 60, 5, 80, 15, PRISM_case_number
  CheckBox 50, 80, 20, 10, "CP", CP
  CheckBox 120, 80, 25, 10, "NCP", NCP
  EditBox 175, 95, 25, 15, Percent
  CheckBox 20, 160, 170, 10, "Check to add CAAD note of documents received", CAAD_checkbox
  CheckBox 15, 255, 120, 10, "Notice of Intent to Enforce UN/UN", Enforce_checkbox
  EditBox 210, 250, 40, 15, amount
  CheckBox 15, 265, 75, 10, "Affidavit of Service ", Aff_Service_checkbox
  EditBox 85, 210, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 290, 50, 15
    CancelButton 240, 290, 50, 15
  Text 10, 10, 50, 10, "Case Number"
  Text 15, 35, 255, 15, "This script will gernerate DORD DOCS F0944, F0659, and F0945 for collection of Unreimbursed and Uninsured Medical and Dental Expenses."
  Text 5, 65, 175, 10, "Check who requested Unreimbursed/Uninsured forms"
  Text 90, 80, 15, 10, "or"
  Text 5, 100, 165, 10, "Enter the PERCENT owed by non requesting party:"
  Text 15, 125, 260, 10, "**************************************************************************************"
  Text 5, 145, 195, 10, "Completed Documents received from Requesting Party (CP)"
  Text 40, 175, 115, 10, "Affidavit of Health Care Expenses"
  Text 40, 185, 145, 10, "Notice to Collect UN Med Exp Req Party"
  Text 40, 195, 110, 10, "Copies of bill, receipts, EOB's"
  Text 5, 240, 165, 10, "Documents to send to Non Requesting party (NCP)"
  Text 140, 255, 65, 10, "Amount Requested"
  Text 20, 215, 60, 10, "Worker Signature"
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
		IF CP = 1 AND NCP = 1 AND Percent = "" THEN err_msg = err_msg & vbNewline & "Percent of Unreimbursed Uninsured Expense must be completed."
		IF CP = 1 AND NCP = 0 AND Percent = "" THEN err_msg = err_msg & vbNewline & "Percent of Unreimbursed Uninsured Expense must be completed."
		IF CP = 0 AND NCP = 1 AND Percent = "" THEN err_msg = err_msg & vbNewline & "Percent of Unreimbursed Uninsured Expense must be completed."
		IF CP = 0 AND NCP = 0 AND Percent <> "" THEN err_msg = err_msg & vbNewline & "You must select either CP or NCP if a percent of un/un is entered."
		IF Enforce_checkbox = 1 and amount = "" THEN err_msg = err_msg & vbNewline & "Please add amount of un/un expenses."
		IF CAAD_checkbox =1 AND worker_signature = "" THEN err_msg = err_msg & vbNewline & "Please sign your CAAD Note."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'END LOOP--------------------------------------

'2nd dialog box for date on aff of service
DIM date, DATE_SERVED_dialog, date_served, confidential_checkbox

IF Aff_Service_checkbox = 1 THEN

BeginDialog DATE_SERVED_dialog, 0, 0, 146, 75, "DATE SERVED"
  EditBox 50, 5, 50, 15, date_served
  CheckBox 10, 30, 125, 10, "check if address is CONFIDENTIAL", confidential_checkbox
  ButtonGroup ButtonPressed
    OkButton 35, 55, 50, 15
    CancelButton 90, 55, 50, 15
  Text 10, 10, 40, 10, "Served on: "
EndDialog



Do
	err_msg = ""
	Dialog DATE_SERVED_dialog
		IF buttonpressed = 0 then stopscript
		IF date_served = "" THEN err_msg = err_msg & vbNewline & "Please enter date you are sending Affidavit of Service."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

Loop until err_msg = ""

END IF


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

	'shift f2, to get to user lables 
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
MsgBox "IMPORTANT!!  IMPORTANT!!" & vbNewline & vbNewline & "First select the correct LEGAL HEADING and press enter, " & vbNewline & "then PRESS OK so script can continue.", vbSystemModal, "Select Legal Heading"



EMWriteScreen "B", 3, 29
transmit

END IF

'ADDS CAAD NOTE
IF CAAD_checkbox = 1 THEN
	CALL navigate_to_PRISM_screen ("CAAD")																					
	PF5
	EMWriteScreen "A", 3, 29
	EMWriteScreen "free", 4, 54
	EMSetCursor 16, 4
'this will add information to the CAAD note of what emc docs sent 
	CALL write_variable_in_CAAD ("CP returned Affidavit of Health Care Expenses, Notice to Collect UN MED   Exp Req Party, and Copies of bills, receipts, EOB's.")
	CALL write_variable_in_CAAD ("Amount requested $" & amount)
	CALL write_variable_in_CAAD(worker_signature)
	transmit
	PF3
END IF

'SENDING DORD to NCP notice of intent to enforce and aff of service F0949
IF  Enforce_checkbox = 1 THEN
	CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0949", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit
	PF14
	PF8
	PF8	

	EMWriteScreen "S", 11, 5
	transmit 
	
	EMWriteScreen(amount), 16, 15
	transmit
	PF3
	EMWriteScreen "m", 3, 29
	transmit
END IF

'DORD aff of service
IF Aff_Service_checkbox = 1 AND confidential_checkbox = 0 THEN
	CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0016", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit
'shift f2, to get to user lables 
	PF14
	EMWriteScreen "u", 20, 14
	transmit 
	PF8
	PF8
	EMWriteScreen "s", 15, 5
	EMWriteScreen "s", 16, 5
	EMWriteScreen "s", 17, 5
	transmit 
	EMWriteScreen "Notice of Intent to Enforce Unreimbursed and/or Uninsured", 16, 15
	transmit 
	EMWriteScreen "Medical/Dental Expenses", 16, 15
	transmit
	EMWriteScreen (date_served), 16, 15
	transmit 
	PF8
	EMWriteScreen "s", 8, 5
	transmit
	EMWriteScreen "N", 16, 15
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit

'''need to select legal heading
MsgBox "IMPORTANT!!  IMPORTANT!!" & vbNewline & vbNewline & "First select the correct LEGAL HEADING and press enter, " & vbNewline & "then PRESS OK so script can continue.", vbSystemModal, "Select Legal Heading"



END IF


IF Aff_Service_checkbox = 1 AND confidential_checkbox = 1 THEN
		CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0016", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit
'shift f2, to get to user lables 
	PF14
	EMWriteScreen "u", 20, 14
	transmit 
	PF8
	PF8
	EMWriteScreen "s", 15, 5
	EMWriteScreen "s", 16, 5
	EMWriteScreen "s", 17, 5
	transmit 
	EMWriteScreen "Notice of Intent to Enforce Unreimbursed and/or Uninsured", 16, 15
	transmit 
	EMWriteScreen "Medical/Dental Expenses", 16, 15
	transmit
	EMWriteScreen (date_served), 16, 15
	transmit 
	PF8
	EMWriteScreen "s", 8, 5
	transmit
	EMWriteScreen "Y", 16, 15
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit

'''need to select legal heading
MsgBox "IMPORTANT!!  IMPORTANT!!" & vbNewline & vbNewline & "First select the correct LEGAL HEADING and press enter, " & vbNewline & "then PRESS OK so script can continue.", vbSystemModal, "Select Legal Heading"



END IF

script_end_procedure("")
