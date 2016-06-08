'option explicit
'GATHERING STATS=================================
name_of_script = "NOTES - INTAKE DOCS RECEIVED.vbs"
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


'DIMMING VARIABLES
DIM beta_agency, row, col, case_number_valid, intake_docs_recd_dialog, paternity_wkst_check, rec_of_parentage_check, prism_case_number, date_recd, app_supp_coll_services_check, app_fee_check, ref_supp_coll_app_check, good_cause_check, role_county_atty_check, aff_arrears_check, waiver_pers_service_check, birth_check, marriage_check, court_order_check, photo_check, insurance_check, ButtonPressed, other_recd, worker_signature   


'THE DIALOG BOX-------------------------------------------------------------------------------------------------------------------

BeginDialog intake_docs_recd_dialog, 0, 0, 341, 200, "Intake Documents Received"
  EditBox 75, 5, 75, 15, prism_case_number
  EditBox 225, 5, 65, 15, date_recd
  CheckBox 15, 40, 165, 10, "Application for Support/Coll Services DHS-1958", app_supp_coll_services_check
  CheckBox 195, 40, 75, 10, "$25 Application Fee", app_fee_check
  CheckBox 15, 55, 155, 10, "Referral to Support/Collections DHS-3163B", ref_supp_coll_app_check
  CheckBox 195, 55, 130, 10, "Client Stmt of Good Cause DHS-2338", good_cause_check
  CheckBox 15, 70, 85, 10, "Role of County Attorney", role_county_atty_check
  CheckBox 195, 70, 105, 10, "Notarized Affidavit of Arrears", aff_arrears_check
  CheckBox 15, 85, 100, 10, "Waiver of Personal Service", waiver_pers_service_check
  CheckBox 195, 85, 50, 10, "Court Order", court_order_check
  CheckBox 15, 100, 105, 10, "Paternity Worksheet/Affidavit", paternity_wkst_check
  CheckBox 195, 100, 125, 10, "Health/Dental Insurance Verification", insurance_check
  CheckBox 15, 115, 100, 10, "Marriage License/Certificate", marriage_check
  CheckBox 195, 115, 95, 10, "Recognition of Parentage", rec_of_parentage_check
  CheckBox 15, 130, 65, 10, "Birth Verification", birth_check
  CheckBox 195, 130, 30, 10, "Photo", photo_check
  EditBox 35, 145, 290, 15, other_recd
  EditBox 75, 175, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 180, 50, 15
    CancelButton 280, 180, 50, 15
  Text 5, 10, 70, 10, "Prism Case Number:"
  Text 185, 10, 40, 15, "Date Rec'd:"
  Text 15, 150, 20, 10, "Other:"
  GroupBox 5, 25, 325, 145, "Documents Rec'd:"
  Text 5, 180, 70, 10, "Sign your CAAD note:"
EndDialog

'THE SCRIPT CODE-------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)


'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

Do
	err_msg = ""
	'Shows dialog, validates that PRISM is up and not timed out, with transmit
	Dialog intake_docs_recd_dialog
	If buttonpressed = 0 then stopscript
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = False THEN err_msg = err_msg & vbNewLine & "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX.  "
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "Sign your CAAD note."
	IF err_msg <> "" THEN
				MsgBox "***NOTICE***" & vbcr & err_msg & vbNewline & vbNewline & "Please resolve for this script to continue."
	END IF
LOOP UNTIL err_msg = ""


'Makes sure you are not passworded out
CALL check_for_PRISM(True)

'Navigates to CAAD and adds the note
CALL navigate_to_PRISM_screen("CAAD")

'Adds new CAAD note
PF5

EMWriteScreen "A", 3, 29

'Writes the CAAD note
EMWriteScreen "FREE", 4, 54                                     'Type of CAAD note
EMWriteScreen "*Intake Documents Received*", 16, 4              'Types "Intake Documents Received" on first line of CAAD note
EMSetCursor 17, 4                                               'Sets the cursor on the next line
IF date_recd <> "" THEN CALL write_bullet_and_variable_in_CAAD("Date Rec'd", date_recd)  'Types in date received on the second lind of CAAD note
IF app_supp_coll_services_check = checked THEN CALL write_variable_in_CAAD("* Application for Support/Coll Services DHS-1958")   'If any of the buttons are checked they will caad note
IF app_fee_check = checked THEN CALL write_variable_in_CAAD("* $25 Application Fee")
IF ref_supp_coll_app_check = checked THEN CALL write_variable_in_CAAD("* Referral to Supp/Coll DHS-3163B")
IF good_cause_check = checked THEN CALL write_variable_in_CAAD("* Client Stmt of Good Cause DHS-2338")
IF role_county_atty_check = checked THEN CALL write_variable_in_CAAD("* Role of the County Attorney")
IF aff_arrears_check = checked THEN CALL write_variable_in_CAAD("* Notarized Affidavit of Arrears")
IF waiver_pers_service_check = checked THEN CALL write_variable_in_CAAD("* Waiver of Personal Service")
IF court_order_check = checked THEN CALL write_variable_in_CAAD("* Court Order")
IF paternity_wkst_check = checked THEN CALL write_variable_in_CAAD("* Paternity Worksheet/Affidavit")
IF insurance_check = checked THEN CALL write_variable_in_CAAD("* Health/Dental Insurance Verification")
IF marriage_check = checked THEN CALL write_variable_in_CAAD(" *Marriage License/Certificate")
IF rec_of_parentage_check = checked THEN CALL write_variable_in_CAAD("* Recognition of Parentage")
IF birth_check = checked THEN CALL write_variable_in_CAAD("* Birth Verification")
IF photo_check = checked THEN CALL write_variable_in_CAAD("* Photo")


IF other_recd <> "" THEN CALL write_bullet_and_variable_in_CAAD("Other", other_recd)  'Other rec'd line

CALL write_variable_in_CAAD(worker_signature)  'Worker signature


'Saves the CAAD note
transmit

'Exits back out of that CAAD note
PF3

script_end_procedure("")          'Stops the script
