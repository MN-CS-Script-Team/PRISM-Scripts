'GATHERING STATS----------------------------------------------------------------------------------------------------

name_of_script = "ACTIONS - COLA.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block 
'-------------------------------------------------------------------------------------------------------------------

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


'connecting to bluezone
EMConnect ""

'to pull up my prism 
EMFocus

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

'brings me to the CAPS screen to auto fill prims case number in dialog
CALL navigate_to_PRISM_screen ("CAPS")
EMReadScreen PRISM_case_number, 13, 4, 8 


'COLAFLOW DIALOG---------------------------------------------------------------------------------------------------------
DIM PRISM_case_number, ButtonPressed, COLAFLOW_Dialog, err_msg

BeginDialog COLAFLOW_Dialog, 0, 0, 116, 60, "COLA Report Default Flow"
  EditBox 10, 20, 95, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 5, 40, 50, 15
    CancelButton 65, 40, 50, 15
  Text 5, 10, 105, 10, "Case Number for Cola Review:"
EndDialog


'THE LOOP for dialog colaflow---------------------------------------------------------------------------------------------- 
Do
	err_msg = ""
	Dialog COLAFLOW_Dialog
		IF buttonpressed = 0 THEN stopscript		'Cancel pressed script ends
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "You must enter a valid Case Number!"
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "No Case Number entered.  Script Ended."
		END IF

Loop until err_msg = ""

'END LOOP for dialog colaflow--------------------------------------------------------------------------------------------

DIM SUOLresult, NCODresult, ARREARSresult, CASEresult, COLAresult 

'suol mn order?
call navigate_to_PRISM_screen("SUOL")
	SUOLresult = msgbox ("Is the charging tied to a MN order?", VbOKCancel)
	IF SUOLresult = vbCancel THEN
		MsgBox "Case may not be appropriate for COLA. Please review case further before continuing. Script Ended."
		StopScript
	End IF

'ncod charging tied to corret order
call navigate_to_PRISM_screen("NCOD")
EMWriteScreen "B", 3, 29
transmit

	NCODresult = msgbox ("Is the current charging tied to the correct order?" & VbNewline & VbNewline & _
 		"Does the Order allow for COLA (Appendix A or language in the Order)?" & VbNewline & VbNewline & _
 		"Check ONBASE." & VbNewline & VbNewline & _ 
 		"Reindex Orders. ", VbOKCancel)
	If NCODresult = vbCancel then
		MsgBox "Case may not be appropriate for COLA. Please review case further before continuing. Script Ended."
		StopScript
	End If

'non standard arrears payment ordered
	ARREARSresult = msgbox ("Is there a non-standard arrears payment ordered (not 20%)?" & VbNewline & VbNewline & _
		"If YES make sure non-accrual is on NCOD and case is coded with 20% overide on SUOD.", VbOKCancel)
	If ARREARSresult = vbCancel then
		MsgBox "Case may not be appropriate for COLA. Please review case further before continuing. Script Ended."
		StopScript
	End If

'check if another case is involed
call navigate_to_PRISM_screen("NCCB")
	CASEresult = msgbox ("Has the charging been moved to/from a related case (foster care or redirect)?", VbOKCancel)
	If CASEresult = vbCancel then
		MsgBox "Case may not be appropriate for COLA. Please review case further before continuing. Script Ended."
		StopScript
	End If

'check cola screen
call navigate_to_PRISM_screen("COLA")
	COLAresult = msgbox ("Review COLA Screen and confirm the information is correct.", VbOKCancel)
	If COLAresult = vbCancel then
		MsgBox "Case may not be appropriate for COLA. Please review case further before continuing. Script Ended."
		StopScript
	End If

FINALResult = msgbox ("Is the COLA okay to run?", VbYesNo)
	If FINALResult = VbNo THEN 
		stopscript
	End If

DIM MN_order, correct, special_arrears, worker_signature, COLACAAD, FINALResult

BeginDialog COLACAAD, 0, 0, 191, 145, "COLA OK TO RUN"
  CheckBox 5, 55, 45, 10, "MN Order.", MN_order
  CheckBox 5, 70, 130, 10, "Charging is tied to the correct Order.", correct
  CheckBox 5, 85, 180, 10, "Non standard arrears collection is loaded correctly.", special_arrears
  EditBox 65, 100, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 125, 50, 15
    CancelButton 135, 125, 50, 15
  Text 50, 10, 80, 10, "IF COLA IS OK TO RUN"
  Text 20, 35, 155, 10, "Check options below to add to your CAAD note."
  Text 5, 105, 60, 10, "Initials for CAAD:"
EndDialog

Do	
		
	err_msg = ""
	Dialog COLACAAD				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		'IF MN_order = 0 AND correct = 0 AND special_arrears = 0 THEN err_msg = err_msg & vbNewline & "You must select something to put in your CAAD note!"
		IF worker_signature = "" THEN MsgBox "Please sign your CAAD Note."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue or press CANCEL to STOP SCRIPT."
		END IF

Loop until err_msg = ""


'bring to CAAD screen to create a CAAD note
	CALL navigate_to_PRISM_screen ("CAAD")																					
	PF5
	EMWriteScreen "A", 3, 29
	EMWriteScreen "free", 4, 54
	EMSetCursor 16, 4

'this will add information to the CAAD note
	CALL write_variable_in_CAAD("~*~CASE REVIEWED FOR COLA - OK TO RUN~*~")
	IF MN_Order = 1 THEN CALL write_variable_in_CAAD ("*MN ORDER.")
	IF correct = 1 THEN CALL write_variable_in_CAAD ("*Charging is tied to the correct Order.")
	IF special_arrears = 1 THEN CALL write_variable_in_CAAD ("*Non Standard arrears collection verified and loaded correctly.")
	CALL write_variable_in_CAAD(worker_signature)
	'transmit
	'PF3

script_end_procedure("")

