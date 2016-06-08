'Gathering stats--------------------------------------------------------------------------------
script_name = "NOTES - NO PAY MONTHS 1 THRU 4.vbs"
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


'DIM variables from dialog; you can include a space and underscore return to start a new line and DIM will read it otherwise it has to be all on one line
'DIM No_Payment_Main_Dialog, ButtonPressed, Case_Number, No_Payment_Reason, NCP_Receiving_PA_dropdownlist, Social_security_benefits_droplist, Month_Dropdownlist, worker_signature, ButtonGroup

'This is intended to have one main dialog box with the worker making a selection of which month they are working and the different steps needed with each month.
'The worker completes the info in the main dialog box and selects the month. The dialog box for that month then opens for worker to complete
'This script will CAAD note and create a worklist reminder if chosen

'MAIN dialog box
BeginDialog No_Payment_Main_Dialog, 0, 0, 271, 225, "No Payment Months 1-4"
  EditBox 60, 10, 75, 15, Case_Number
  Text 5, 35, 215, 15, "NCP was called for status update regarding no payment received."
  EditBox 130, 60, 135, 15, No_Payment_Reason
  DropListBox 135, 95, 130, 15, "Select one:"+chr(9)+"Yes receiving public assistance"+chr(9)+"No public assistance case", NCP_Receiving_PA_dropdownlist
  DropListBox 135, 125, 130, 15, "Select one:"+chr(9)+"Yes receiving Social Security benefits"+chr(9)+"No Social Security benefits issued", Social_security_benefits_droplist
  DropListBox 65, 180, 105, 15, "Select one:"+chr(9)+"Month ONE"+chr(9)+"Month TWO"+chr(9)+"Month THREE"+chr(9)+"Month FOUR", Month_Dropdownlist
  EditBox 60, 205, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 205, 50, 15
    CancelButton 215, 205, 50, 15
  Text 5, 10, 50, 15, "Case Number"
  Text 5, 60, 120, 15, "Reason no payment has been made:"
  Text 5, 80, 125, 30, "Confirmed via MAXIS, NCP is receiving public assistance and coded NCDE panel 2:"
  Text 5, 120, 125, 20, "Confirmed via SSTD and SSSD, NCP is receiving Social Security benefits:"
  Text 5, 210, 50, 10, "Worker Name:"
  Text 25, 160, 200, 15, "Select the appropriate month below for additional questions:"
EndDialog

'Month ONE dialog DIM-----------------------------------------------------------------------------------------------
'DIM No_Payment_Month_One_Dialog, NCP_DL_Loaded_CheckBox, SUDL_DL_Suppressed_checkbox, CRB_Suppressed_Checkbox, New_HIRE_checkbox, Create_worklist_CheckBox

'dialog box for Month ONE
BeginDialog No_Payment_Month_One_Dialog, 0, 0, 276, 135, "No Payment - Month One"
  Text 5, 10, 265, 10, " NCP was called for status update regarding no payment received - Month ONE."
  CheckBox 30, 25, 145, 15, "Is NCP driver's license loaded in NCLD?", NCP_DL_Loaded_CheckBox
  CheckBox 30, 40, 155, 15, "Is DL suppressed in SUDL appropriately?", SUDL_DL_Suppressed_checkbox
  CheckBox 30, 55, 155, 15, "Is CRB suppressed in SUDL appropriately?", CRB_Suppressed_Checkbox
  CheckBox 5, 75, 205, 15, "Confirmed all employer information is current in NCLD", New_HIRE_checkbox
  CheckBox 5, 90, 180, 15, "Create a worker worklist note for 30 days from today.", Create_worklist_CheckBox
  ButtonGroup ButtonPressed
    OkButton 160, 115, 50, 15
    CancelButton 220, 115, 50, 15
EndDialog


'Month TWO dialog DIM---------------------------------------------------------------------------------------------------
'DIM No_Payment_Month_Two_Dialog, Reached_NCP_Checkbox, Sent_NO_PAYMENT_Letter_CheckBox, Sent_NCPWFC_Letter_CheckBox, Request_Manual_Requests_Checkbox, Check_LOID_Checkbox

'dialog box for Month TWO
BeginDialog No_Payment_Month_Two_Dialog, 0, 0, 281, 150, "No Payment - Month Two"
  Text 25, 10, 255, 15, "NCP was called for status update regarding no payment received Month TWO."
  CheckBox 5, 25, 250, 15, "Call attempt to NCP was successful. No payment letter sent.", Reached_NCP_Checkbox
  Text 5, 45, 80, 15, "NCP was NOT reached:"
  CheckBox 90, 40, 85, 15, "Sent no payment letter", Sent_NO_PAYMENT_Letter_CheckBox
  CheckBox 5, 60, 180, 15, "Requested manual requests through NCMR for DLI", Request_Manual_Requests_Checkbox
  CheckBox 5, 75, 180, 15, "Checked LOID for additional information", Check_LOID_Checkbox
  CheckBox 5, 100, 180, 15, "Create a worker worklist note for 30 days from today.", Create_worklist_CheckBox
  ButtonGroup ButtonPressed
    OkButton 170, 130, 50, 15
    CancelButton 225, 130, 50, 15
EndDialog


'Month THREE Dialog DIM---------------------------------------------------------------------------------------------------
'DIM No_Payment_Month_Three_Dialog, Sent_Pay2_letter_checkbox, Moving_to_Contempt_Dropdownlist


'Dialog box for Month THREE
BeginDialog No_Payment_Month_Three_Dialog, 0, 0, 281, 100, "No Payment - Month Three"
  Text 5, 5, 270, 10, "NCP was called for status update regarding no payment received - Month THREE."
  CheckBox 5, 20, 95, 10, "Sent Pay 2 Letter to NCP", Sent_Pay2_letter_checkbox
  Text 5, 40, 150, 10, "Case appears to be moving toward contempt action:"
  DropListBox 160, 40, 115, 15, "Select one:"+chr(9)+"Yes payment history created"+chr(9)+"No payment history created", Moving_to_Contempt_Dropdownlist
  CheckBox 5, 60, 180, 15, "Create a worker worklist note for 30 days from today.", Create_worklist_CheckBox
  ButtonGroup ButtonPressed
    OkButton 165, 80, 50, 15
    CancelButton 225, 80, 50, 15
EndDialog

'Month FOUR Dialog DIM---------------------------------------------------------------------------------------------------
'DIM No_Payment_Month_Four_Dialog, Check_ENFL_checkbox, Contempt_Checkbox, prohibit_contempt, Call_CP_Checkbox, CP_Info



'Dialog box for Month FOUR
BeginDialog No_Payment_Month_Four_Dialog, 0, 0, 296, 225, "No Payment - Month FOUR"
  Text 15, 5, 270, 10, "NCP was called for status update regarding no payment received - Month FOUR."
  Text 5, 35, 260, 20, "NOTE: If DL notice was sent more than 30 days ago, case is ready for contempt or the contempt list."
  CheckBox 5, 60, 120, 15, "Checked ENFL for driver's license", Check_ENFL_checkbox
  CheckBox 5, 80, 155, 15, "Case is ready for contempt or contempt list", Contempt_Checkbox
  Text 5, 105, 155, 15, "CSO is aware of factor(s) that prohibit contempt:"
  EditBox 165, 105, 125, 15, prohibit_contempt
  CheckBox 5, 125, 130, 15, "Called CP for additional information", Call_CP_Checkbox
  Text 5, 150, 100, 15, "Additional information from CP"
  EditBox 110, 150, 180, 15, CP_Info
  CheckBox 5, 185, 180, 15, "Create a worker worklist note for 30 days from today.", Create_worklist_CheckBox
  ButtonGroup ButtonPressed
    OkButton 185, 205, 50, 15
    CancelButton 240, 205, 50, 15
EndDialog



'Connecting to BlueZone
EMConnect ""
'Check_for_PRISM

'Case Note dialog - need to do the DO loops right after running dialog
DO
	err_msg = ""
	DIALOG No_Payment_Main_Dialog
		IF ButtonPressed = 0 THEN StopScript
		IF case_number = "" THEN err_msg = err_msg & vbCr & "* Please enter a case number."
		IF no_payment_reason = "" THEN err_msg = err_msg & vbCr & "* Please enter a no-payment reason."
		IF NCP_receiving_PA_dropdownlist = "Select one:" THEN err_msg = err_msg & vbCr & "* Please confirm receipt of public assistance."
		IF social_security_benefits_droplist = "Select one:" THEN err_msg = err_msg & vbCr & "* Please confirm receipt of Social Security benefits."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your CAAD note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Month One
IF Month_Dropdownlist = "Month ONE" THEN

		DIALOG No_Payment_Month_ONE_Dialog
		IF ButtonPressed = 0 THEN StopScript
'Month TWO
ELSEIF Month_Dropdownlist = "Month TWO" THEN

		DIALOG No_Payment_Month_TWO_Dialog
		IF ButtonPressed = 0 THEN StopScript

'Month THREE

ELSEIF Month_Dropdownlist = "Month THREE" THEN 

		DIALOG No_Payment_Month_THREE_Dialog
		IF ButtonPressed = 0 THEN StopScript

'Month FOUR
ELSEIF Month_Dropdownlist = "Month FOUR" THEN
		DIALOG No_Payment_Month_FOUR_Dialog
		IF ButtonPressed = 0 THEN StopScript

END IF



'check for an autofill case number function - otherwise this works with CSO entering case number.
'EMReadScreen PRISM_case_number, 13, 20, 8

'Navigate to CAAD
EMWriteScreen "CAAD", 21, 18
EMSendKey "<enter>"
EMWaitReady 0, 0 

'transmit to PRISM
EMSendKey "<PF5>"
EMWaitReady 0, 0 


EMWriteScreen "T0055", 4, 54

'Set Cursor on first line of CAAD note entry and start entering responses from MAIN Dialog box 
EMSetCursor 16, 4
CALL write_variable_in_CAAD ("NCP stated no payments made due to " & NCP_No_Payment_Reason)
IF NCP_DL_Loaded_CheckBox = 1 THEN CALL write_variable_in_CAAD("*  NCP's DL is loaded in NCLD.")
IF SUDL_DL_Suppressed_Checkbox = 1 THEN CALL write_variable_in_CAAD("*  DL is suppressed in SUDL appropriately.")
IF CRB_Suppressed_Checkbox = 1 THEN CALL write_variable_in_CAAD("*  CRB suppressed in SUDL appropriately.") 
IF New_Hire_checkbox = 1 THEN CALL write_variable_in_CAAD("*  Confirmed all employer information is current in NCLD.")
CALL write_variable_in_CAAD ("Public Assistance confirmed via MAXIS: " & NCP_Receiving_PA_dropdownlist)
CALL write_variable_in_CAAD ("SS benefits confirmed via SSTD and SSSD: " & Social_security_benefits_droplist)

'writes CAAD note for Month ONE
IF Month_Dropdownlist = "Month ONE" THEN

	CALL write_variable_in_CAAD("**No Payment - Month ONE**")
	IF NCP_DL_Loaded_CheckBox = 1 THEN CALL write_variable_in_CAAD("NCP driver's license is loaded in NCLD")
	IF SUDL_DL_Suppressed_checkbox = 1 THEN CAll write_variable_in_CAAD("DL is suppressed in SUDL")
	IF CRB_Suppressed_Checkbox = 1 THEN CAll write_variable_in_CAAD("CRB is suppressed in SUDL")
	IF New_HIRE_Checkbox = 1 THEN CAll write_variable_in_CAAD("Employer information is current in NEW HIRE")

END IF


'writes CAAD note for Month TWO
IF Month_Dropdownlist = "Month TWO" THEN

	CALL write_variable_in_CAAD("**No Payment - Month TWO**")
	IF Reached_NCP_Checkbox = 1 THEN CALL write_variable_in_CAAD("Call attempt to NCP was successful.  No payment letter sent")
	IF Sent_NO_PAYMENT_Letter_CheckBox = 1 THEN CAll write_variable_in_CAAD("NCP was not reached: No payment letter sent")
	IF Request_Manual_Requests_Checkbox = 1 THEN CAll write_variable_in_CAAD("Requested manual requests through NCMR for DLI")
	IF Check_LOID_Checkbox = 1 THEN CAll write_variable_in_CAAD("LOID was checked for additional information")

END IF


'writes CAAD note for Month THREE
IF Month_Dropdownlist = "Month THREE" THEN
	CALL write_variable_in_CAAD("**No Payment - Month THREE**")
	CALL write_variable_in_CAAD ("Case appears to be moving toward contempt action: " & Moving_to_Contempt_Dropdownlist)

END IF


'writes CAAD note for Month FOUR
IF Month_Dropdownlist = "Month FOUR" THEN

	CALL write_variable_in_CAAD("**No Payment - Month THREE**")
	IF Check_ENFL_checkbox = 1 THEN CAll write_variable_in_CAAD("ENFL was checked for driver's license")
	IF Contempt_Checkbox = 1 THEN CAll write_variable_in_CAAD("Case is ready for contempt or contempt list")
	CALL write_variable_in_CAAD("The following factor(s) prohibit contempt:  " & prohibit_contempt)
	IF Call_CP_Checkbox = 1 THEN CAll write_variable_in_CAAD("CP provided additional information:  " & CP_Info)

END IF 
CALL write_variable_in_CAAD ("----" & worker_signature)

EMWriteScreen "A", 3, 29
transmit

IF Create_worklist_CheckBox = 0 THEN script_end_procedure("Success!!")

'****************REMINDER SET-UP WORKLIST REMINDER ******************
CALL navigate_to_PRISM_screen("CAWD")

'Creating a new worklist
PF5

EMWriteScreen "FREE", 4, 37
EMWriteScreen "Review for payment from NCP", 10, 4
EMWriteScreen "30", 17, 52
transmit


'RETURNS TO MAIN SCREEN "CAST" 
CALL navigate_to_PRISM_screen("CAST")
transmit

script_end_procedure("Success!!")
