'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CSENET INFO.vbs"
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

'DIMMING VARIABLES-------------------------------------------------------------------------------------------------------------------------------------

DIM prism_case_number, csenet_total, csenet_info_dialog, ButtonPressed, write_new_line_in_CAAD, csenet_sent_recd, reason_code_line, row, col, beta_agency, worker_signature, case_number_valid, csenet_dateline, csenet_line_01, csenet_line_02, csenet_line_03, csenet_line_04, csenet_line_05



'THE DIALOG-------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog csenet_info_dialog, 0, 0, 216, 90, "CSENET Info"
  EditBox 85, 20, 60, 15, prism_case_number
  EditBox 85, 45, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 105, 70, 50, 15
    CancelButton 160, 70, 50, 15
  Text 10, 50, 70, 10, "Sign you CAAD note:"
  Text 10, 25, 70, 15, "Prism Case Number:"
  Text 10, 5, 235, 10, "Make sure INTD message is open before running this script."
EndDialog


'THE SCRIPT-----------------------------------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if


'Makes sure you are not passworded out
CALL check_for_PRISM(True)

'The script will not run unless the CAAD note is signed and there is a valid prism case number
DO
	DO
		Dialog csenet_info_dialog
		IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
		IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                   'If worker sig is blank, message box pops saying you must sign caad note
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
	LOOP UNTIL case_number_valid = True
LOOP UNTIL worker_signature <> ""                                                                  'Will keep popping up until worker signs note


'Reads the contents of the CSENET for CAAD noting
EMReadScreen reason_code_line, 45, 13, 14
EMReadScreen csenet_sent_recd, 1, 14, 61
EMReadScreen csenet_line_01, 80, 15, 2
EMReadScreen csenet_line_02, 80, 16, 2
EMReadScreen csenet_line_03, 80, 17, 2
EMReadScreen csenet_line_04, 80, 18, 2
EMReadScreen csenet_line_05, 80, 19, 2

csenet_line_01 = replace(csenet_line_01, "_", "")
csenet_line_02 = replace(csenet_line_02, "_", "")
csenet_line_03 = replace(csenet_line_03, "_", "")
csenet_line_04 = replace(csenet_line_04, "_", "")
csenet_line_05 = replace(csenet_line_05, "_", "")

csenet_total = csenet_line_01 & " " & csenet_line_02 & " " & csenet_line_03 & " " & csenet_line_04 & " " & csenet_line_05

'Navigates to CAAD and adds the note
CALL navigate_to_PRISM_screen("CAAD")


'Adds new CAAD note
PF5
     
EMSetCursor 16, 4
CALL write_variable_in_CAAD("##CSENET INFO##")                                      'Writes CSENET INFO in title
                                                                      
CALL write_bullet_and_variable_in_CAAD("CSESNET sent/rcd", csenet_sent_recd)         'Writes CSENET Sent/Recd and Date/Time
CALL write_bullet_and_variable_in_CAAD("Reason Code", reason_code_line)             'Writes in the reason code
CALL write_bullet_and_variable_in_CAAD("CSENET Comments", csenet_total)             'Writes CSENET Comments
CALL write_variable_in_CAAD("---")				                              'Writes Worker Signature
CALL write_variable_in_CAAD(worker_signature)


EMWriteScreen "A", 3, 29                                                          'Writes A to add the new caad note

'Writes the CAAD note type
EMWriteScreen "T0111", 4, 54

EMWriteScreen PRISM_case_number, 4, 8

'Saves the CAAD note
transmit

'Exits back out of that CAAD note
PF3

script_end_procedure("")             'Stops the script
