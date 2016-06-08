option explicit

'STATS GATHERING----------------------------------------------------------------------------------------------------
Dim name_of_script, start_time
name_of_script = "NOTES - COURT ORDER REQUEST.vbs"
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
DIM beta_agency, row, col, case_number_valid, Court_Order_Request_Dialog, prism_case_number, date_court_order_requested, requested_via_droplistbox, requested_from, court_order_number, create_worklist_checkbox, worker_signature, order_type, ButtonPressed


'THE DIALOG----------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog Court_Order_Request_Dialog, 0, 0, 406, 95, "Court Order Request"
  EditBox 80, 5, 70, 15, prism_case_number
  EditBox 290, 5, 65, 15, date_court_order_requested
  ComboBox 100, 30, 115, 15, "Click here to enter county name"+chr(9)+"CP"+chr(9)+"NCP", requested_from
  DropListBox 290, 30, 85, 15, "Select one..."+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Telephone"+chr(9)+"Mail"+chr(9)+"SIR Email"+chr(9)+"Inter-Office", requested_via_droplistbox
  EditBox 80, 50, 85, 15, court_order_number
  EditBox 290, 50, 100, 15, order_type
  EditBox 80, 75, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 295, 75, 50, 15
    CancelButton 350, 75, 50, 15
  Text 235, 30, 50, 10, "Requested Via:"
  Text 5, 10, 70, 10, "Prism Case Number:"
  Text 5, 80, 70, 10, "Sign your CAAD note:"
  Text 190, 10, 95, 10, "Date Court Order Requested:"
  Text 245, 55, 40, 10, "Order Type:"
  Text 15, 55, 65, 10, "Court File Number:"
  Text 5, 30, 90, 20, "Requested From: (or type in County name)"
EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

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

'The script will not run unless the CAAD note is signed, and you are in a valid Prism case, and makes requested from field mandatory
DO
	DO
		DO
			Dialog court_order_request_dialog
			IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
			IF worker_signature = "" THEN MsgBox "You must sign your CAAD note!"                   'If worker sig is blank, message box pops saying you must sign caad note
			CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
		LOOP UNTIL worker_signature <> ""                                                            'Will keep popping up until worker signs note
		IF requested_from = "Select one..." THEN MsgBox "You must complete 'Requested From field'"   'Makes this field mandatory
	LOOP UNTIL requested_from <> "Select one..."
LOOP UNTIL case_number_valid = TRUE


'Navigates to CAAD and adds note
CALL navigate_to_PRISM_screen("CAAD")

'Adds a new CAAD note
PF5

EMWriteScreen "A", 3, 29


'Writes the CAAD NOTE
EMWriteScreen "B0170", 4, 54         'Type of Caad note
EMWriteScreen date_court_order_requested, 4, 37
EMSetCursor 16, 4
CALL write_bullet_and_variable_in_CAAD("Date Court Order Requested", date_court_order_requested)   'types date court order requested info
CALL write_bullet_and_variable_in_CAAD("Requested From", requested_from)                           'types requested from info
CALL write_bullet_and_variable_in_CAAD("Requested Via", requested_via_droplistbox)			   'types requested via info
CALL write_bullet_and_variable_in_CAAD("Court File Number", court_order_number)                    'types court file number info
CALL write_bullet_and_variable_in_CAAD("Order Type", order_type)                                   'types order type info
CALL write_variable_in_CAAD(worker_signature)                                                      'types worker signature


'Saves the CAAD note
transmit

'Exits back out of that CAAD note
PF3

'Navigates to CAWT and creates worklist
CALL navigate_to_PRISM_screen("CAWT")

'Adds a new worklist
PF5

'Puts the A in the Action part
EMWriteScreen "A", 3, 30

'Writes the Worklist
EMWriteScreen "FREE", 4, 37
'Writes note in CAWT
EMWriteScreen "Did Court Order come in? See CAAD notes for request details.", 10, 4

script_end_procedure("")
