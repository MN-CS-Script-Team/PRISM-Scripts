option explicit

DIM beta_agency, row, col

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


'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen prism_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if


DIM service_of_process, prism_case_number, invoice_number, dollar_amount, service_date, legal_action, person_served, service_checkbox, pay_yes_checkbox, signature, buttonpressed

'Calling dialog details for the Service of Process---------------------------------------------------------------------
BeginDialog service_of_process, 0, 0, 226, 215, "Service of Process"
  EditBox 50, 5, 65, 15, prism_case_number
  EditBox 50, 25, 65, 15, invoice_number
  EditBox 50, 45, 65, 15, dollar_amount
  EditBox 70, 65, 55, 15, service_date
  ComboBox 70, 95, 115, 15, "select one, or type action...."+chr(9)+"Contempt"+chr(9)+"Establishment"+chr(9)+"Paternity", legal_action
  ComboBox 70, 125, 115, 15, "select one, or type person served....."+chr(9)+"ALF"+chr(9)+"CP"+chr(9)+"NCP", person_served
  CheckBox 10, 150, 125, 10, "Check if service was successfull", service_checkbox
  CheckBox 10, 165, 115, 10, "Check if invoice is ok to pay", pay_yes_checkbox
  EditBox 65, 180, 35, 15, signature
  ButtonGroup ButtonPressed
    OkButton 115, 195, 50, 15
    CancelButton 170, 195, 50, 15
  Text 10, 10, 30, 10, "Case #:"
  Text 10, 30, 35, 10, "Invoice #:"
  Text 10, 50, 35, 10, "$ Amount:"
  Text 10, 70, 55, 10, "Date of service:"
  Text 10, 85, 55, 25, "Legal Action: (choose one or type action)"
  Text 10, 115, 60, 25, "Person served: (choose from list or fill in name)"
  Text 10, 185, 55, 10, "Worker initials:"
EndDialog


'Connecting to Bluezone
EMConnect ""			

'checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

DO																'inserting the loop so that the date and signature are required fields (to start the loop type DO)
	Dialog service_of_process											'open the dialog box itself
	IF ButtonPressed = 0 THEN StopScript									'if cancel button is pressed, the script will stop running
	IF signature = "" THEN MSGbox "Please sign your CAAD Note"						'if the signature is blank pop up a message box
	IF IsDate(service_date) = False THEN MsgBox "You must enter a valid date"			'makes sure the date field is a valid date
LOOP UNTIL signature <>"" and IsDate(service_date) = TRUE							'tells the loop to keep running until the signature field is filled in and the date is valid.  (if you have a Do stmt, you must have a LOOP UNTIL stmt)
		

'go to CAAD
CALL Navigate_to_PRISM_screen ("CAAD")										'goes to the CAAD screen
PF5																'F5 to add a note
EMWritescreen "A", 3, 29												'put the A on the action line

'writes info from dialog into caad
EMWritescreen "FREE", 4, 54												'types free on caad code: line
EMWritescreen "Service Notes:", 16, 4										'types title of the free caad on the first line of the note
EMSetCursor 17, 4														'puts the cursor on the very next line to be ready to enter the info

call write_bullet_and_variable_in_CAAD("invoice #",invoice_number)
call write_bullet_and_variable_in_CAAD("$",dollar_amount)
call write_bullet_and_variable_in_CAAD("service date", service_date)
call write_bullet_and_variable_in_CAAD("Legal action", legal_action)
call write_bullet_and_variable_in_CAAD("person served", person_served)  
If service_checkbox = 1 then call write_variable_in_CAAD("service was successful")
If service_checkbox = 0 then call write_variable_in_CAAD("service was not successful")
If pay_yes_checkbox = 1 then call write_variable_in_CAAD("Invoice is OK to pay")
If pay_yes_checkbox = 0 then call write_variable_in_CAAD("Do Not pay invoice")
call write_variable_in_CAAD(signature)
transmit
PF3

script_end_procedure("")                                                                     	'stopping the script






