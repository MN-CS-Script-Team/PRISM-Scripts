'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - E-FILING.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 60
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

BeginDialog efiling_dialog, 0, 0, 186, 200, "E-Filing"
  EditBox 80, 5, 100, 15, prism_case_number
  EditBox 85, 30, 95, 15, action_type
  EditBox 80, 50, 100, 15, doc_efiled
  DropListBox 70, 75, 95, 15, "Select One..."+chr(9)+"Submitted"+chr(9)+"Accepted", efile_status_dropdown
  CheckBox 10, 100, 140, 10, "Check here to add a follow-up worklist", worklist_checkbox
  EditBox 75, 115, 105, 15, envelope_number
  EditBox 75, 135, 105, 15, eservice
  EditBox 75, 155, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 180, 50, 15
    CancelButton 130, 180, 50, 15
  Text 10, 80, 50, 10, "E-Filing Status:"
  Text 10, 35, 70, 10, "Type of Legal Action:"
  Text 10, 55, 65, 10, "Documents E-Filed:"
  Text 10, 120, 65, 10, "Envelope Number:"
  Text 5, 10, 70, 10, "PRISM Case Number:"
  Text 10, 140, 60, 10, "E-Service Details:"
  Text 10, 160, 60, 10, "Worker Signature:"
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
	Dialog efiling_dialog
	cancel_confirmation
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF action_type = "" THEN err_msg = err_msg & vbNewline & "You must enter a type of legal action!"
	IF doc_efiled = "" THEN err_msg = err_msg & vbNewline & "You must enter the type of documents you E-Filed!"
	IF efile_status_dropdown = "Select One..." THEN err_msg = err_msg & vbNewline & "You must select an E-Filing Status!" 
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'NAVIGATES TO CAAD
CALL navigate_to_PRISM_screen("CAAD")

'ENTERING CASE NUMBER
CALL enter_PRISM_case_number(PRISM_case_number, 20, 8)

'ADDS NEW CAAD NOTE WITH FREE CAAD CODE
PF5 
EMWritescreen "FREE", 4, 54

'SETS THE CURSOR
EMSetCursor 16, 4

'WRITES THE CAAD NOTE
IF efile_status_dropdown = "Submitted" THEN CALL write_variable_in_CAAD("E-Filing Status: Documents Submitted")
IF efile_status_dropdown = "Accepted" THEN CALL write_variable_in_CAAD("E-Filing Status: Documents Accepted")
CALL write_bullet_and_variable_in_CAAD("Type of Action", action_type)
CALL write_bullet_and_variable_in_CAAD("Documents E-Filed", doc_efiled)
CALL write_bullet_and_variable_in_CAAD("Envelope Number", envelope_number)
CALL write_bullet_and_variable_in_CAAD("E-Service Details", eservice)
CALL write_variable_in_CAAD(worker_signature)
transmit

'SENDS MSG BOX TO WORKER TO REMIND THEM TO UPDATE LEHD IF NECESSARY
IF efile_status_dropdown = "Accepted" THEN Msgbox "***REMINDER***" & vbNewline & "Do you need to update the LEHD screen with new court file number?"

'ADDS A WORKLIST IF THE CHECKBOX TO ADD ONE IS CHECKED
IF worklist_checkbox = CHECKED THEN 
	CALL navigate_to_PRISM_screen("CAWT")
	PF5
	EMWritescreen "FREE", 4, 37

	'SETS THE CURSOR AND STARTS THE WORKLIST
	IF efile_status_dropdown = "Submitted" THEN EMWritescreen "E-Filing Status: Documents Submitted", 10, 4
	IF efile_status_dropdown = "Accepted" THEN EMWritescreen "E-Filing Status: Documents Accepted", 10, 4
	EMSetCursor 11,4
	IF envelope_number <> "" THEN CALL write_bullet_and_variable_in_CAAD("Envelope Number", envelope_number)
END IF

'REMINDS WORKER TO FINISH AND SAVE THEIR WORKLIST
IF worklist_checkbox = CHECKED THEN 
	script_end_procedure("Please finish and save your worklist")
ELSE
	script_end_procedure("")
END IF
