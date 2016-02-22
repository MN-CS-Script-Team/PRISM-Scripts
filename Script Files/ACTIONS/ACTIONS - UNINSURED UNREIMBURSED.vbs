'Option Explicit 'this has to be on the top, always
Option Explicit

'this is a function document
DIM beta_agency 'remember to add

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO                                                                          'Declares variables to be good to option explicit users
If beta_agency = "" then                                              'For scriptwriters only
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then                 'For beta agencies and testers
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else                                                                                                                        'For most users
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")                                                               'Creates an object to get a URL
req.open "GET", url, False                                                                                                                                            'Attempts to open the URL
req.send                                                                                                                                                                                                              'Sends request
If req.Status = 200 Then                                                                                                                                                '200 means great success
                Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
                Execute req.responseText                                                                                                                          'Executes the script code
ELSE                                                                                                                                                                                                                       'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
                MsgBox                "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
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
'this is where the copy and paste from functions library ended



'DIALOGS---------------------------------------------------------------------------
DIM UnUn_Dialog, PRISM_case_number, CP, NCP, Percent, err_msg, ButtonPressed, case_number_is_valid

BeginDialog UnUn_Dialog, 0, 0, 291, 145, "Unreimbursed Uninsured Docs"
  EditBox 60, 45, 90, 15, PRISM_case_number
  CheckBox 50, 85, 20, 10, "CP", CP
  CheckBox 120, 85, 25, 10, "NCP", NCP
  EditBox 175, 100, 25, 15, Percent
  ButtonGroup ButtonPressed
    OkButton 180, 125, 50, 15
    CancelButton 235, 125, 50, 15
  Text 25, 10, 240, 15, "This script will gernerate DORD DOCS F0944, F0659, and F0945 for collection of Unreimbursed and Uninsured Medical and Dental Expenses."
  Text 5, 50, 50, 10, "Case Number"
  Text 5, 70, 175, 10, "Check who requested Unreimbursed/Uninsured forms"
  Text 90, 85, 15, 10, "or"
  Text 5, 105, 165, 10, "Enter the PERCENT owed by non requesting party:"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'to pull up my prism 
EMFocus

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
		IF Percent = "" THEN err_msg = err_msg & vbNewline & "Percent of Unreimbursed Uninsured Expense must be completed."
		'IF both cp box and ncp box blank
		IF CP = 0 AND NCP = 0 THEN err_msg = vbNewline & "Either CP or NCP must be selected."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'END LOOP--------------------------------------


'brings me to caad and creates DORD doc for NCP
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

	EMSendKey "<enter>" 
	CALL write_variable_in_CAAD (Percent)
	EMSendKey "<enter>"
	PF3
	EMWriteScreen "M", 3, 29
	transmit

END IF

'brings me to caad and creates DORD doc for CP
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
	EMSendKey "<enter>" 
	CALL write_variable_in_CAAD (Percent)
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit

End IF

'''need to select legal heading''''''''''''''''''''''''''''HELP HELP is there a better way'''''''''''''''''''''''''''''''''''''''''HELP

MsgBox ( "IMPORTANT!!  IMPORTANT!!" & vbNewline & vbNewline & "First select the correct LEGAL HEADING and press enter, " & vbNewline & "then PRESS OK so script can continue." )


EMWriteScreen "B", 3, 29
transmit


script_end_procedure("")

