'option explicit -- COMMENTED OUT PER VKC REQUEST
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - ROP DETAIL.vbs"
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


'DIMMING variables
DIM beta_agency, row, col, case_number_valid, prism_case_number, worker_signature, ButtonPressed, cp_dob, ncp_dob, cp_rop_signed, ncp_rop_signed, rop_dialog, child_name, rop_completed_at, cp_turned_18_date, ncp_turned_18_date


'THE DIALOG--------------------------------------------------------------------------------------------------

BeginDialog Rop_Dialog, 0, 0, 246, 220, "ROP Detail"
  EditBox 80, 5, 85, 15, prism_case_number
  EditBox 80, 25, 85, 15, child_name
  ComboBox 80, 45, 125, 15, "Select One or Type In..."+chr(9)+"County"+chr(9)+"DHS"+chr(9)+"Hospital"+chr(9)+"MN Dept of Health", rop_completed_at
  EditBox 55, 90, 50, 15, cp_dob
  EditBox 175, 90, 50, 15, cp_rop_signed
  EditBox 55, 140, 50, 15, ncp_dob
  EditBox 175, 140, 50, 15, ncp_rop_signed
  EditBox 80, 180, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 200, 50, 15
    CancelButton 190, 200, 50, 15
  GroupBox 5, 125, 235, 45, "Non-Custodial Parent Info"
  Text 5, 10, 75, 10, "PRISM Case Number:"
  Text 120, 95, 55, 10, "ROP Sign Date:"
  Text 5, 185, 70, 10, "Sign your CAAD note:"
  Text 10, 145, 40, 10, "NCP's DOB:"
  Text 30, 30, 50, 10, "Child's Name:"
  Text 15, 95, 35, 10, "CP's DOB:"
  Text 15, 50, 65, 10, "ROP completed at:"
  GroupBox 5, 75, 235, 40, "Custodial Parent Info"
  Text 120, 145, 55, 10, "ROP Sign Date:"
EndDialog


'THE SCRIPT-------------------------------------------------------------------------------------------------

'Connects to Bluezone
EMConnect ""                    

'Brings Bluezone to the front
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


'The script will not run unless the CAAD note is signed, has a valid prism number, and ROP completed at is completed
DO	
	DO
		DO
			Dialog rop_dialog
			IF ButtonPressed = 0 THEN StopScript		                                            'Pressing Cancel stops the script
			IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                        'If worker signature field is blank, message box will pop up to instruct worker to sign note
			CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
		LOOP UNTIL worker_signature <> ""                                                                                           'Will keep popping up until worker signs note
		IF rop_completed_at = "Select One or Type In..." THEN MsgBox "You must make a selection in the 'ROP Completed at' field"    'Makes this field mandatory
	LOOP UNTIL rop_completed_at <> "Select One or Type In..."
LOOP UNTIL case_number_valid = TRUE

'Calculating math difference between parties DOB and ROP signed date
cp_turned_18_date = DateAdd("yyyy", 18, cp_dob)
ncp_turned_18_date = DateAdd("yyyy", 18, ncp_dob)


'Navigates to CAAD and adds note
CALL navigate_to_PRISM_screen("CAAD")


'Adds a new CAAD note
PF5
EMWritescreen "A", 3, 29


'Writes the CAAD NOTE
EMSetCursor 4, 54
EMWriteScreen "FREE", 4, 54     'Type of Caad note
EMWriteScreen "*ROP DETAIL*", 16, 4
EMSetCursor 17, 4
CALL write_bullet_and_variable_in_CAAD("ROP Details for child", child_name)
CALL write_bullet_and_variable_in_CAAD("ROP Completed At", rop_completed_at)
CALL write_bullet_and_variable_in_CAAD("CP ROP Signed on", cp_rop_signed)
CALL write_bullet_and_variable_in_CAAD("NCP ROP Signed on", ncp_rop_signed)
IF DateDiff("d", cp_turned_18_date, cp_rop_signed) < 0 THEN CALL write_variable_in_CAAD("* CP was under 18 when ROP was signed")
IF DateDiff("d", ncp_turned_18_date, ncp_rop_signed) < 0 THEN CALL write_variable_in_CAAD("* NCP was under 18 when ROP was signed")
CALL write_variable_in_CAAD(worker_signature)
transmit
PF3	


'Creates a message box if either party was under 18 when ROP was signed
IF DateDiff("d", cp_turned_18_date, cp_rop_signed) < 0 THEN MsgBox "CP was under 18 when ROP was signed!!!"
IF DateDiff("d", ncp_turned_18_date, ncp_rop_signed) <0 THEN MsgBox "NCP was under 18 when ROP was signed!!!"


script_end_procedure("")
