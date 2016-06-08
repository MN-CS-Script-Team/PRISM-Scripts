'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - ARREARS MGMT REVIEW.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog arrears_mgmt_dialog, 0, 0, 386, 246, "Arrears Mgmt Review"
  DropListBox 100, 30, 270, 20, "E9852 Reviewed for Arrears Mgmt - Approval Requested"+chr(9)+"E9851 Reviewed for Arrears Mgmt - No Action"+chr(9)+"E9853 Reviewed for Arrears Mgmt - More Information Needed", CAAD_type
  EditBox 170, 0, 80, 20, PRISM_case_number
  EditBox 60, 100, 320, 20, date_ranges
  EditBox 90, 130, 290, 20, details
  EditBox 110, 160, 260, 20, arrears_mgmt_amount
  CheckBox 30, 50, 260, 10, "Check here if arrears mgmt is for CMS while NCP was also a recipient of MA", CMS_check
  CheckBox 30, 60, 250, 10, "Check here if arrears mgmt is for charging while NCP rec'd cash assistance", Cash_PA_check
  CheckBox 30, 180, 330, 10, "Check here if your request for arrears mgmt includes suspension of PA interest charging", suspend_interest_check
  CheckBox 30, 70, 250, 10, "Check here if arrears mgmt is for charging while NCP was incarcerated", incarcerated_check
  EditBox 90, 210, 70, 20, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 210, 50, 20
    CancelButton 310, 210, 50, 20
  Text 10, 30, 90, 10, "Please select CAAD note:"
  Text 10, 10, 160, 10, "PRISM case number (XXXXXXXXXX-XX format):"
  Text 10, 100, 50, 10, "Date ranges:"
  Text 10, 130, 80, 20, "Other details about this arrears mgmt review:"
  Text 10, 160, 100, 10, "Total amount of arrears mgmt:"
  Text 10, 210, 70, 10, "Sign your case note: "
  CheckBox 30, 80, 270, 10, "Check here if arrears mgmt is for other circumstances", other_check
  CheckBox 30, 190, 330, 10, "Check here if your request for arrears mgmt includes $1 PA forgiveness for every $1 payment", dollar_for_dollar_check
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
		dialog arrears_mgmt_dialog
		IF buttonpressed = 0 THEN stopscript
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = FALSE THEN MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	LOOP UNTIL case_number_valid = TRUE
			
	CALL navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	IF case_activity_detail <> "Case Activity Detail" THEN MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
LOOP UNTIL case_activity_detail = "Case Activity Detail"


IF CAAD_type = "E9852 Reviewed for Arrears Mgmt - Approval Requested" THEN CAAD_code = "E9852"
IF CAAD_type = "E9851 Reviewed for Arrears Mgmt - No Action" THEN CAAD_code = "E9851"
IF CAAD_type = "E9853 Reviewed for Arrears Mgmt - More Information Needed" THEN CAAD_code = "E9853"

'Writing the case note
EMWriteScreen CAAD_code, 4, 54				

EMSetCursor 16, 4 								'Because the PRISM case note functions require the cursor to start here
if details <> "" THEN CALL write_bullet_and_variable_in_CAAD("Arrears Mgmt Review Details", details)
if date_ranges <> "" THEN CALL write_bullet_and_variable_in_CAAD("Date ranges", date_ranges)
if arrears_mgmt_amount <> "" THEN CALL write_bullet_and_variable_in_CAAD("Amount requested", arrears_mgmt_amount)
if dollar_for_dollar_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt request includes $1 PA forgiveness for every $1 paid, if approved.")
if suspend_interest_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt request includes suspension of PA interest charging, if approved.")
if CMS_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt requested because CMS was charged while NCP was also a recipient of MA.") 
if Cash_PA_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt requested because NCP received cash public assistance.") 
IF incarcerated_check = 1 THEN CALL write_variable_in_CAAD("* Arrears Mgmt requested because NCP was incarcerated.")
if other_check = 1 THEN CALL write_variable_in_CAAD("* Arrears Mgmt requested ")
CALL write_variable_in_CAAD("---")
CALL write_variable_in_CAAD(worker_signature)

script_end_procedure("")

