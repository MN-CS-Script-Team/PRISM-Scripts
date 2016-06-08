'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - FRAUD REFERRAL.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------------

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

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------------
'THE DIALOG---------------------------------------------------

BeginDialog fraud_referral_dialog, 0, 0, 301, 280, "Fraud Referral"
  EditBox 80, 5, 85, 15, prism_case_number
  EditBox 130, 30, 160, 15, ref_reason
  EditBox 115, 75, 145, 15, case_num
  EditBox 115, 95, 145, 15, cp_fw
  CheckBox 30, 130, 25, 15, "CCC", ccc_checkbox
  CheckBox 65, 130, 30, 15, "DWP", dwp_checkbox
  CheckBox 105, 130, 25, 15, "FCC", fcc_checkbox
  CheckBox 140, 130, 30, 15, "MAO", ma_checkbox
  CheckBox 175, 130, 25, 15, "MFP", mfp_checkbox
  DropListBox 70, 165, 125, 15, "Select One..."+chr(9)+"Email"+chr(9)+"Telephone Call"+chr(9)+"Fax"+chr(9)+"Mail", ref_droplistbox
  EditBox 70, 185, 125, 15, sent_to
  CheckBox 10, 215, 205, 10, "Check here to create a worklist to check on referral status in", worklist_checkbox
  EditBox 215, 210, 20, 15, days_out
  EditBox 75, 235, 135, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 260, 50, 15
    CancelButton 220, 260, 50, 15
  Text 10, 10, 70, 10, "Prism Case Number:"
  Text 10, 240, 60, 10, "Worker Signature:"
  Text 35, 100, 75, 10, "CP's Financial Worker:"
  Text 10, 35, 120, 10, "Reason for sending Fraud Referral:"
  Text 10, 165, 60, 10, "Referral Sent Via:"
  Text 20, 80, 95, 10, "MAXIS/METS Case Number:"
  GroupBox 5, 55, 285, 95, "Financial Assistance Case Details"
  Text 240, 215, 20, 10, "days"
  Text 10, 190, 60, 10, "Referral Sent To:"
  Text 20, 115, 70, 10, "PA Program Open:"
EndDialog


'Connects to BLUEZONE
EMConnect ""

'MAKES SURE YOU ARE NOT PASSWORDED OUT
CALL check_for_PRISM(True)

'AUTO POPULATES PRISM CASE NUMBER INTO DIALOG
CALL PRISM_case_number_finder(PRISM_case_number)

'MAKES THINGS MANDATORY
DO
	err_msg = ""
	Dialog fraud_referral_dialog
	cancel_confirmation
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF ref_reason = "" THEN err_msg = err_msg & vbNewline & "You must enter a reason for the fraud referral!"
	IF case_num = "" THEN err_msg = err_msg & vbNewline & "You must enter a MAXIS/METS case number!"
	IF ref_droplistbox = "Select One..." THEN err_msg = err_msg & vbNewline & "You must select how you sent the Referral!" 
	IF sent_to = "" THEN err_msg = err_msg & vbNewline & "You must enter who the Referral was sent to!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'NAVIGATES TO CAAD
CALL navigate_to_PRISM_screen("CAAD")

'ENTERING CASE NUMBER
CALL enter_PRISM_case_number(PRISM_case_number, 20, 8)


'MAKING THE PA PROGRAM OPEN NEATER
IF ccc_checkbox = checked THEN prog_open = prog_open & "CCC,"
IF dwp_checkbox = checked THEN prog_open = prog_open & " DWP,"
IF fcc_checkbox = checked THEN prog_open = prog_open & " FCC,"
IF ma_checkbox = checked THEN prog_open = prog_open & " MAO,"
IF mfp_checkbox = checked THEN prog_open = prog_open & " MFP,"
prog_open = trim(prog_open)
IF right(prog_open, 1) = "," THEN prog_open = left(prog_open, len(prog_open) - 1)
 

'ADDS NEW CAAD NOTE WITH FREE CAAD CODE
PF5 
EMWritescreen "M0010", 4, 54

'SETS THE CURSOR
EMSetCursor 16, 4

'WRITES THE CAAD NOTE
CALL write_bullet_and_variable_in_CAAD("Reason for sending Fraud Referral", ref_reason)
CALL write_bullet_and_variable_in_CAAD("MAXIS/METS Case Number", case_num)
CALL write_bullet_and_variable_in_CAAD("CP's financial worker", cp_fw)
CALL write_bullet_and_variable_in_CAAD("PA Program Open", prog_open)
CALL write_bullet_and_variable_in_CAAD("Referral Sent Via", ref_droplistbox)
CALL write_bullet_and_variable_in_CAAD("Referral Sent to", sent_to)
CALL write_variable_in_CAAD(worker_signature)
transmit


'ADDS A WORKLIST IF THE CHECKBOX TO ADD ONE IS CHECKED
IF worklist_checkbox = CHECKED THEN 
	CALL navigate_to_PRISM_screen("CAWT")
	PF5
	EMWritescreen "FREE", 4, 37

	'SETS THE CURSOR AND STARTS THE WORKLIST
	EMSetCursor 10, 4
	EMWriteScreen "Check status of Fraud Referral made", 10, 4
	EMSetCursor 11, 4
	CALL write_bullet_and_variable_in_CAAD("Reason for Referral", ref_reason)
	'EMSetCursor 17, 52
	EMWritescreen days_out, 17, 52

END IF

script_end_procedure("")
