'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLIENT CONTACT.vbs"
start_time = timer

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
date_of_contact = date & ""	'defaults to today

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
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

'DIM contact_dialog, contact_type_cp, contact_type_ncp, contact_type_other, verified_ID_check, PRISM_case_number, phone_number, time_contact_was_made, issue, actions_taken, verifs_needed, special_instructions_for_client, left_generic_message_check, worker_signature, ButtonPressed, create_mainframe_friendly_date


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog contact_dialog, 0, 0, 381, 295, "Client contact"
  DropListBox 80, 15, 260, 15, ""+chr(9)+"T0050 PHONE CALL TO CP"+chr(9)+"T0051 PHONE CALL FR CP"+chr(9)+"T0052 PHONE CALL RET TO CP"+chr(9)+"T0053 PHONE CALL RET FR CP"+chr(9)+"T0054 PHONE CALL ATMPT TO RET TO CP"+chr(9)+"T0093 CONTACT WITH CP SPOUSE"+chr(9)+"T0101 PHONE CONTACT CP'S ATTORNEY"+chr(9)+"T0201 CONTACT WITH CP EMPLOYER"+chr(9)+"M3910 INTERVIEW WITH CP", contact_type_CP
  DropListBox 80, 35, 260, 15, ""+chr(9)+"T0055 PHONE CALL TO NCP"+chr(9)+"T0056 PHONE CALL FR NCP"+chr(9)+"T0057 PHONE CALL RET TO NCP"+chr(9)+"T0058 PHONE CALL RET FR NCP"+chr(9)+"T0059 PHONE CALL ATMPT TO RET TO NCP"+chr(9)+"T0060 PHONE CALL TO NCP EMP"+chr(9)+"T0061 PHONE CALL FROM NCP EMP"+chr(9)+"T0062 PHONE CALL RET TO NCP EMP"+chr(9)+"T0063 PHONE CALL RET FR NCP EMP"+chr(9)+"T0064 PHONE CALL ATMPT RET TO NCP EMP"+chr(9)+"T0065 PHONE CALL TO NCP AY"+chr(9)+"T0066 PHONE CALL FR NCP AY"+chr(9)+"T0067 PHONE CALL RET TO NCP AY"+chr(9)+"T0068 PHONE CALL RET FR NCP AY"+chr(9)+"T0069 PHONE CALL ATMPT RET TO NCP AY"+chr(9)+"T0092 CONTACT WITH NCP SPOUSE"+chr(9)+"M3911 INTERVIEW WITH NCP", contact_type_NCP
  DropListBox 80, 55, 260, 15, ""+chr(9)+"M0410 CONTACT WITH CCC WORKER"+chr(9)+"T0070 PHONE CALL/OTHER"+chr(9)+"T0074 CONTACT WITH STATE HELP DESK"+chr(9)+"T0075 CONTACT WITH HEALTH CARRIER"+chr(9)+"T0080 CONTACT WITH COURT ADMINISTRATOR"+chr(9)+"T0085 CONTACT WITH LAW ENFORCEMENT"+chr(9)+"T0087 CONTACT WITH PROBATION OFFICER"+chr(9)+"T0090 CONTACT WITH NCP/CP UNION"+chr(9)+"T0095 CONTACT WITH SOCIAL WORKER"+chr(9)+"T0098 CONTACT WITH WORKER FROM ANOTHER MN COUNTY"+chr(9)+"T0100 PHONE CONTACT WITH OTHER STATE'S CENTRAL REGISTRY"+chr(9)+"T0102 PHONE CONTACT COUNTY ATTORNEY"+chr(9)+"T0103 PHONE CONTACT WITH OTHER STATE WORKER"+chr(9)+"T0104 PHONE CONTACT WITH FINANCIAL WORKER"+chr(9)+"T0105 PHONE CONTACT WITH CSPC"+chr(9)+"T0111 CONTACT WITH OTHER STATE AGENCY"+chr(9)+"T1107 CONTACT WITH MN VITAL RECORDS", contact_type_other
  EditBox 165, 80, 80, 15, PRISM_case_number
  EditBox 310, 80, 70, 15, date_of_contact
  EditBox 95, 110, 60, 15, phone_number
  EditBox 285, 110, 85, 15, time_contact_was_made
  EditBox 55, 135, 325, 15, issue
  EditBox 55, 155, 325, 15, actions_taken
  EditBox 65, 185, 310, 15, verifs_needed
  EditBox 120, 205, 255, 15, special_instructions_for_client
  CheckBox 5, 230, 150, 10, "Check here if you verified ID.", verified_ID_check
  CheckBox 5, 245, 230, 10, "Check here if you left a generic message requesting they return call.", left_generic_message_check
  EditBox 310, 255, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 275, 50, 15
    CancelButton 330, 275, 50, 15
  Text 15, 20, 55, 10, "CP contact type:"
  Text 15, 40, 60, 10, "NCP contact type:"
  Text 15, 60, 65, 10, "Other contact type:"
  Text 5, 85, 160, 10, "PRISM case number (XXXXXXXXXX-XX format):"
  Text 250, 85, 55, 10, "Date of contact:"
  GroupBox 5, 100, 370, 30, "Optional contact info:"
  Text 40, 115, 50, 10, "Phone number: "
  Text 195, 115, 85, 10, "Time contact was made: "
  Text 5, 140, 50, 10, "Issue/subject: "
  Text 5, 160, 50, 10, "Actions taken: "
  GroupBox 5, 175, 375, 50, "Helpful optional case info"
  Text 15, 190, 50, 10, "Verifs needed: "
  Text 15, 210, 100, 10, "Special instructions for client:"
  Text 235, 260, 70, 10, "Sign your case note: "
  GroupBox 5, 5, 370, 70, "Select one contact type from this group, based on CAAD note requirement"
EndDialog


'DIM row, col, EMSearch, EMReadScreen

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
IF row <> 0 THEN
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	IF isnumeric(left(PRISM_case_number, 10)) = FALSE OR isnumeric(right(PRISM_case_number, 2)) = FALSE THEN PRISM_case_number = ""
END IF

'Shows dialog, then navigates to CAAD. It will validate the PRISM case number using the custom function.
DO
	DO
		DO
			DO
				dialog contact_dialog
				IF buttonpressed = 0 THEN stopscript
				CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
				IF case_number_valid = FALSE THEN MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
			LOOP UNTIL case_number_valid = TRUE
			IF ((contact_type_CP <> "" and contact_type_NCP = "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP <> "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP = "" and contact_type_other <> "")) = False then MsgBox("Please select one (and only one) of the contact type options.")
		LOOP UNTIL (contact_type_CP <> "" and contact_type_NCP = "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP <> "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP = "" and contact_type_other <> "")
		IF isdate(date_of_contact) = FALSE THEN MsgBox "You must put a valid date in as the date of contact. Please try again."
	LOOP UNTIL isdate(date_of_contact) = TRUE
	CALL navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	IF case_activity_detail <> "Case Activity Detail" THEN MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
LOOP UNTIL case_activity_detail = "Case Activity Detail"

'Determining which of the three contact_type editboxes contains info, and then making that the "contact_type" variable
IF contact_type_CP <> "" and contact_type_NCP = "" and contact_type_other = "" THEN contact_type = contact_type_CP
IF contact_type_CP = "" and contact_type_NCP <> "" and contact_type_other = "" THEN contact_type = contact_type_NCP
IF contact_type_CP = "" and contact_type_NCP = "" and contact_type_other <> "" THEN contact_type = contact_type_other

'Writing the case note
'Contact date is now included in body of text.  Script does not change Activty Date
EMWriteScreen left(contact_type, 5), 4, 54				'The contact type (only need the left 5 characters)

EMSetCursor 16, 4 								'Because the PRISM case note functions require the cursor to start here
IF issue <> "" THEN CALL write_bullet_and_variable_in_CAAD("Issue/subject", issue)
CALL write_bullet_and_variable_in_CAAD("Date of Contact", date_of_contact & ", " & time_contact_was_made)
IF verified_ID_check = 1 THEN CALL write_variable_in_CAAD("* Verified ID.")
IF actions_taken <> "" THEN CALL write_bullet_and_variable_in_CAAD("Actions taken", actions_taken)
IF verifs_needed <> "" THEN CALL write_bullet_and_variable_in_CAAD("Verifs needed", verifs_needed)
IF special_instructions_for_client <> "" THEN CALL write_bullet_and_variable_in_CAAD("Special Instructions for Client", special_instructions_for_client)
IF case_status <> "" THEN CALL write_bullet_and_variable_in_CAAD("Case status", case_status)
IF left_generic_message_check = 1 THEN CALL write_variable_in_CAAD("* Left client a generic message requesting a return call.")
IF phone_number <> "" THEN CALL write_bullet_and_variable_in_CAAD("Phone number", phone_number)
CALL write_variable_in_CAAD("---")
CALL write_variable_in_CAAD(worker_signature)

script_end_procedure("")
