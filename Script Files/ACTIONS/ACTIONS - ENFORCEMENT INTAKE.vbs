'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = "ACTIONS - INTAKE.vbs"
'start_time = timer

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

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
checked = 1
unchecked = 0



'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog CS_intake_dialog, 0, 0, 371, 315, "CS intake dialog"
  CheckBox 15, 30, 145, 10, "Case Opening - Welcome Letter", NCP_welcome_ltr_check
  CheckBox 15, 45, 140, 10, "Court Order Summary", ncp_court_order_summary_check
  CheckBox 15, 60, 50, 10, "DORD F0999 - PIN Notice", NCP_PIN_Notice_Check
  CheckBox 15, 75, 150, 10, "DORD F0924 - Health Insurance Verification", NCP_health_ins_verif_check
  CheckBox 15, 90, 120, 10, "Notice of Arrears Reported", arrears_reported_check
  CheckBox 50, 115, 65, 10, "DORD F0100", dord_F0100_check
  CheckBox 50, 145, 65, 10, "DORD F0109", dord_F0109_check
  CheckBox 50, 175, 60, 10, "DORD F0107", dord_F0107_check
  CheckBox 15, 220, 125, 10, "Set File Location to QC 30", qc_30_file_loc_check
  CheckBox 15, 235, 115, 10, "Set File Location to SAFETY", safety_file_loc_check
  CheckBox 195, 30, 130, 10, "Case Opening - Welcome Letter", CP_welcome_ltr_check
  CheckBox 195, 45, 115, 10, "CP New Order Summary", CP_new_order_summary_check
  CheckBox 195, 60, 50, 10, "DORD F0999 - PIN Notice", CP_PIN_Notice_check
  CheckBox 195, 75, 155, 10, "DORD F0924 - Health Insurance Verification", CP_health_ins_verif_check
  CheckBox 195, 90, 130, 10, "Child Care Verification", child_care_verif_check
  CheckBox 195, 105, 125, 10, "CP Statement of Arrears Letter", CP_Stmt_of_Arrears_check
  CheckBox 200, 135, 105, 10, "10 day tickler to call NCP", t_10_day_tickler_check
  CheckBox 200, 150, 110, 10, "30 day tickler to load arrears", t_30_day_to_load_arrears_check
  CheckBox 200, 165, 105, 10, "30 day case review", t_30_day_case_review_check
  EditBox 210, 175, 140, 15, t_30_day_cawd_txt
  CheckBox 200, 195, 105, 10, "60 day case review", t_60_day_case_review_check
  EditBox 210, 205, 140, 15, t_60_day_cawd_txt
  EditBox 240, 235, 110, 15, worker_name
  EditBox 240, 255, 110, 15, worker_phone
  EditBox 90, 295, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 290, 50, 20
    CancelButton 275, 290, 50, 20
  Text 15, 295, 70, 10, "Sign your CAAD note:"
  Text 35, 105, 30, 10, "NPA"
  Text 35, 135, 95, 10, "MFIP, DWP, CCA"
  Text 35, 165, 90, 10, "MA only"
  Text 185, 255, 50, 10, "Worker phone:"
  Text 185, 235, 50, 10, "Worker name:"
  GroupBox 5, 205, 170, 50, "File Location on CAST"
  GroupBox 5, 15, 170, 180, "Letters to NCP"
  GroupBox 185, 15, 170, 105, "Letters to CP"
  Text 5, 0, 325, 15, "Enforcement Intake Script - your selections appear in a E0001 CAAD"
  GroupBox 185, 125, 170, 105, "CAWD notes to add"
  EditBox 15, 275, 120, 15, add_caad_txt
  Text 5, 260, 85, 10, "Additional CAAD note text"
EndDialog

'CUSTOM FUNCTION***************************************************************************************************************


FUNCTION send_dord_doc(recipient, dord_doc)
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen dord_doc, 6, 36
	EMWriteScreen recipient, 11, 51
	transmit
END FUNCTION
	
'This is a custom function to fix data that we are reading from PRISM that includes underscores.  The parameter is a string for the 
'variable to be searched.  The function searches the variable and removes underscores.  Then, the fix case function is called to format
'the string in the correct case.  Finally, the data is trimmed to remove any excess spaces.	
FUNCTION fix_read_data (search_string) 
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
END FUNCTION

' This is a custom function to change the format of a participant name.  The parameter is a string with the 
' client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".  
FUNCTION change_client_name_to_FML(client_name)
	client_name = trim(client_name)
	length = len(client_name)
	position = InStr(client_name, ", ")
	last_name = Left(client_name, position-1)
	first_name = Right(client_name, length-position-1)	
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
END FUNCTION

'This is a custom function to change the file location on the CAST screen
FUNCTION set_file_loc_on_CAST(new_file_location)
	call navigate_to_PRISM_screen("CAST")
	EMWriteScreen "M", 3, 29
	EMWriteScreen new_file_location, 14, 17
	transmit
END FUNCTION

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
call navigate_to_PRISM_screen("CAPS")

'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it

'Getting worker info for case note
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
EMReadScreen worker_phone, 12, 8, 35
PF3

Do
	EMReadScreen child_name_on_CAPS, 30, PRISM_row, 16	'reading name
	child_name_on_CAPS = trim(child_name_on_CAPS)		'removing spaces from beginning and end
	EMReadScreen child_DOB_on_CAPS, 10, PRISM_row, 64	'reading DOB
	If child_name_on_CAPS <> "" then CAPS_kids = CAPS_kids & child_name_on_CAPS & " (DOB: " & child_DOB_on_CAPS & ")" & chr(13) 		'If there's a name, add to the CAPS_kids variable
	PRISM_row = PRISM_row + 1					'increase the PRISM row
	If PRISM_row = 21 then						'If we're on row 21, go to the next page
		PF8
		PRISM_row = 18
	End if
Loop until child_name_on_CAPS = ""

'Cleaning up worker info
worker_name = trim(worker_name)
call fix_case(worker_name, 1)
worker_name = change_client_name_to_FML(worker_name)

'Get information to pull into documents
EMReadScreen NCP_MCI, 10, 8, 11 
EMReadScreen CP_MCI, 10, 4, 8 	
	
'NCP Name
call navigate_to_PRISM_screen("NCDE")
EMWriteScreen NCP_MCI, 4, 7
EMReadScreen NCP_F, 12, 8, 34
EMReadScreen NCP_M, 12, 8, 56
EMReadScreen NCP_L, 17, 8, 8
EMReadScreen NCP_S, 3, 8, 74
NCP_name = fix_read_data(NCP_F) & " " & fix_read_data(NCP_M) & " " & fix_read_data(NCP_L)	
If trim(NCP_S) <> "" then NCP_name = NCP_Name & " " & ucase(fix_read_data(NCP_S))
NCP_name = trim(NCP_name)
'NCP Address
'Navigating to NCDD to pull address info
call navigate_to_PRISM_screen("NCDD")
EMReadScreen ncp_street_line_1, 30, 15, 11
EMReadScreen ncp_street_line_2, 30, 16, 11
EMReadScreen ncp_city_state_zip, 49, 17, 11

'Cleaning up address info
ncp_street_line_1 = replace(ncp_street_line_1, "_", "")
call fix_case(ncp_street_line_1, 1)
ncp_street_line_2 = replace(ncp_street_line_2, "_", "")
call fix_case(ncp_street_line_2, 1)
if trim (ncp_street_line_2) <> "" then
	ncp_address = ncp_street_line_1 & chr(13) & ncp_street_line_2
else
	ncp_address = ncp_street_line_1
end if
ncp_city_state_zip = replace(replace(replace(ncp_city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(ncp_city_state_zip, 2)


'CP Name											
call navigate_to_PRISM_screen("CPDE")
EMWriteScreen CP_MCI, 4, 7
EMReadScreen CP_F, 12, 8, 34
EMReadScreen CP_M, 12, 8, 56
EMReadScreen CP_L, 17, 8, 8
EMReadScreen CP_S, 3, 8, 74
CP_name = fix_read_data(CP_F) & " " & fix_read_data(CP_M) & " " & fix_read_data(CP_L)	
If trim(CP_S) <> "" then CP_name = CP_Name & " " & ucase(fix_read_data(CP_S))
CP_name = trim(CP_Name)

'CP Address
'Navigating to CPDD to pull address info
call navigate_to_PRISM_screen("CPDD")
EMReadScreen cp_street_line_1, 30, 15, 11
EMReadScreen cp_street_line_2, 30, 16, 11
EMReadScreen cp_city_state_zip, 49, 17, 11

'Cleaning up address info
cp_street_line_1 = replace(cp_street_line_1, "_", "")
call fix_case(cp_street_line_1, 1)
cp_street_line_2 = replace(cp_street_line_2, "_", "")
if trim (cp_street_line_2) <> "" then
	cp_address = cp_street_line_1 & chr(13) & cp_street_line_2
else
	cp_address = cp_street_line_1
end if
call fix_case(cp_street_line_2, 1)
cp_city_state_zip = replace(replace(replace(cp_city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(cp_city_state_zip, 2)

'FINANCIAL SUMMARY
call navigate_to_PRISM_screen("CAFS")
EMReadScreen total_arrears, 14, 12, 64
EMReadScreen total_due, 14, 14, 25
call navigate_to_PRISM_screen("NCOL")
EMWriteScreen "CCH", 20, 39
transmit
EMWriteScreen "S", 9, 3
transmit
EMSetCursor 12, 55
PF1
EMReadScreen order_date, 8, 11, 56
PF3
call navigate_to_PRISM_screen("NCOL")
EMWriteScreen "CCH", 20, 39
transmit
EMReadScreen CCH_amount, 9, 9, 36
EMWriteScreen "CCC", 20, 39
transmit
EMReadScreen CCC_amount, 9, 9, 36
EMWriteScreen "CMI", 20, 39
transmit
EMReadScreen CMI_amount, 9, 9, 36
EMWriteScreen "CMS", 20, 39
transmit
EMReadScreen CMS_amount, 9, 9, 36
EMWriteScreen "JCH", 20, 39
transmit
EMReadScreen JCH_amount, 9, 9, 36

if inStr(Cstr(CCH_amount), "Data") <= 0 then
	Obligation = Obligation & vbCr  & cstr(ccur(CCH_amount)) & " per month ongoing basic support"
end if
if instr(Cstr(CCC_amount), "Data") <= 0 then
	Obligation = Obligation & vbCR & cstr(ccur(CCC_amount)) & " per month child care support"
end if
if inStr(Cstr(CMI_amount), "Data") <= 0 then
	Obligation = Obligation & vbCR & cstr(ccur(CMI_amount)) & " per month medical insurance contribution"
end if
if instr(Cstr(CMS_amount), "Data") <= 0 then
	Obligation = Obligation & vbCR & cstr(ccur(CMS_amount)) & " per month medical support"
end if
if instr(Cstr(JCH_amount), "Data") <= 0 then
	Obligation = Obligation & vbCR & cstr(ccur(JCH_amount)) & " per month toward past due support and/or arrears totaling " & cstr(ccur(total_arrears))
Obligation = Obligation & vbCR & vbCR & "Total monthly support due: " & cstr(ccur(total_due))
end if

'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
Do
	Dialog CS_intake_dialog
	If buttonpressed = 0 then stopscript
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


'Creating the Word application object (if any of the Word options are selected), and making it visible 
If _
	NCP_welcome_ltr_check = checked or _
	ncp_court_order_summary_check = checked or _
	arrears_reported_check = checked or _
	CP_welcome_ltr_check = checked or _
	CP_Stmt_of_Arrears_check = checked or _
	child_care_verif_check = checked or _
	CP_new_order_summary_check = checked then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End if

'NCP Welcome Letter
If NCP_welcome_ltr_check = checked then
'	set objDoc = objWord.Documents.Add("E:\Enforcement Script\NCP Case opening- Welcome Letter.dotm")
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\NCP Case opening- Welcome Letter.dotm")
	With objDoc
		.FormFields("NCPName").Result = NCP_name
		.FormFields("NCPAddress").Result = ncp_address
		.FormFields("NCPCSZ").Result = ncp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CPName").Result = CP_name
	'	.FormFields("OrderDate").Result = order_date
	'	.FormFields("Obligations").Result = Obligations
		.FormFields("NCPMCI").Result = NCP_MCI
		.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'NCP Court Order Summary
If ncp_court_order_summary_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\Court Order Summary Notice.dotm")
	With objDoc
		.FormFields("NCPName").Result = NCP_name
		.FormFields("NCPAddress").Result = ncp_address
		.FormFields("NCPCSZ").Result = ncp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CPName").Result = CP_name
	'	.FormFields("Obligations").Result = Obligations
		.FormFields("NCPMCI1").Result = NCP_MCI
		.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'Arrears Reported
If arrears_reported_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\NCP Arrears Reported Ltr.dotm")
	With objDoc
		.FormFields("NCPName").Result = NCP_name
		.FormFields("NCPName1").Result = NCP_name
		.FormFields("NCPAddress").Result = NCP_address
		.FormFields("NCPCSZ").Result = Ncp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CPName").Result = CP_name
		.FormFields("CPName1").Result = CP_name
		.FormFields("Date30").Result = dateadd("d", date, 30)
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'CP Welcome Letter
If CP_welcome_ltr_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\CP Case Opening - Welcome Letter.dotm")
	With objDoc
		.FormFields("CPName").Result = CP_name
		.FormFields("CPAddress").Result = CP_address
		.FormFields("CPCSZ").Result = cp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
	'	.FormFields("CaseNumber1").Result = PRISM_case_number
		.FormFields("NCPName").Result = NCP_name
		.FormFields("NCPName1").Result = NCP_name
	'	.FormFields("Obligations").Result = Obligations
		.FormFields("CPMCI").Result = CP_MCI
		'.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'CP Statment of Arrears
If CP_Stmt_of_Arrears_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\CP Stmt of Support Cover Letter.dotm")
	With objDoc
		.FormFields("CPName").Result = CP_name
		.FormFields("CPAddress").Result = CP_address
		.FormFields("CPCSZ").Result = cp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CaseNumber1").Result = PRISM_case_number
		.FormFields("NCPName").Result = NCP_name
		.FormFields("NCPName1").Result = NCP_name
		.FormFields("NCPName2").Result = NCP_name
		'.FormFields("Obligations").Result = Obligations
		'.FormFields("NCPMCI").Result = NCP_MCI
		'.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'Child Care Verification
If child_care_verif_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\Childcare Verification Letter.dotm")
	With objDoc
		.FormFields("CPName").Result = CP_name
		.FormFields("CPName1").Result = CP_name
		.FormFields("CPName2").Result = CP_name
		.FormFields("CPAddress").Result = CP_address
		.FormFields("CPCSZ").Result = cp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CaseNumber1").Result = PRISM_case_number
		.FormFields("ReturnDate").Result = dateadd("d", date, 10)
		'.FormFields("Obligations").Result = Obligations
		'.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'CP New Order Summary
If CP_new_order_summary_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Enforcement Script\CP New Order Summary.dotm")
	With objDoc
		.FormFields("CPName").Result = CP_name
		.FormFields("CPName1").Result = CP_name
		.FormFields("CPAddress").Result = CP_address
		.FormFields("CPCSZ").Result = cp_city_state_zip
		.FormFields("CaseNumber").Result = PRISM_case_number
		.FormFields("CPMCI").Result = CP_MCI
		.FormFields("NCPName").Result = NCP_name
		.FormFields("NCPName1").Result = NCP_name
		.FormFields("NCPName2").Result = NCP_name
		.FormFields("NCPName3").Result = NCP_name
	'	.FormFields("Obligations").Result = Obligation
		.FormFields("CaseWorkerName").Result = worker_name
		.FormFields("CaseWorkerTitle").Result = "Child Support Officer"
		.FormFields("CaseWorkerPhone").Result = worker_phone
	End With
End if

'If F0999 is indicated on the dialog then it navigates to DORD to send it.
If NCP_PIN_Notice_Check = checked then 'Send PIN Notice
	call send_dord_doc("NCP", "F0999")
End if

'If F0924 is indicated on the dialog then it navigates to DORD to send it.
If NCP_health_ins_verif_check = checked then 
	call send_dord_doc("NCP", "F0924") 
End if

'If F0100 is indicated on the dialog then it navigates to DORD to send it.
If dord_F0100_check = checked then
	call send_dord_doc("NCP", "F0100")
End if

'If F0109 is indicated on the dialog then it navigates to DORD to send it.
If dord_F0109_check = checked then 
	call send_dord_doc("NCP", "F0109")
End if

'If F0107 is indicated on the dialog then it navigates to DORD to send it.
If dord_F0107_check = checked then
	call send_dord_doc("NCP", "F0107")
End if
'If F0924 is indicated on the dialog then it navigates to DORD to send it.
If CP_health_ins_verif_check = checked then
	call send_dord_doc("CPP", "F0924")
End if
'If F0999 is indicated on the dialog then it navigates to DORD to send it.
If CP_PIN_Notice_check = checked then
	call send_dord_doc("CPP", "F0999")
End if

'************************Change File Location on Cast

If qc_30_file_loc_check = checked then
	set_file_loc_on_CAST("QC 30")
End if


if safety_file_loc_check = checked then
	set_file_loc_on_CAST("Safety")
End if


'**************************Add worklists

If t_10_day_tickler_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** Call NCP to answer any questions NCP has about case setup.", 10, 4
	EMWriteScreen dateadd("d", date, 10), 17, 21
	transmit
End if
If t_30_day_to_load_arrears_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "Load arrears?", 10, 4
	EMWriteScreen dateadd("d", date, 30), 17, 21
	transmit
End if
If t_30_day_case_review_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "30 Day Case Review", 10, 4
	EMWriteScreen t_30_day_cawd_txt, 11, 4
	EMWriteScreen dateadd("d", date, 30), 17, 21
	transmit
End if
If t_60_day_case_review_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "60 Day Case Review", 10, 4
	EMWriteScreen t_60_day_cawd_txt, 11, 4
	EMWriteScreen dateadd("d", date, 60), 17, 21
	transmit
End if

'**********************************




'Going to CAAD, adding a new note
call navigate_to_PRISM_screen("CAAD")
EMWriteScreen "A", 8, 5
transmit
EMReadScreen case_activity_detail, 20, 2, 29
If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")


'Setting the type
EMWriteScreen "E0001", 4, 54


'Setting cursor in write area and writing note details
EMSetCursor 16, 4
call write_new_line_in_PRISM_case_note("* The following documents were sent:")
	If NCP_welcome_ltr_check = checked then call write_new_line_in_PRISM_case_note("    * Case Opening - Welcome letter to NCP")
	If ncp_court_order_summary_check = checked then call write_new_line_in_PRISM_case_note("    * Court Order Summary to NCP")
	If NCP_PIN_Notice_check = checked then call write_new_line_in_PRISM_case_note("    * F0999 - PIN Notice to NCP")
	If NCP_health_ins_verif_check = checked then call write_new_line_in_PRISM_case_note("    * F0924 - Health Insurance Verification to NCP")
	If arrears_reported_check = checked then call write_new_line_in_PRISM_case_note("    * Notice of Arrears Reported to NCP")
	If dord_F0100_check = checked then call write_new_line_in_PRISM_case_note("    * F0100 sent to NCP")
	If dord_F0109_check = checked then call write_new_line_in_PRISM_case_note("    * F0109 sent to NCP")
	If dord_F0107_check = checked then call write_new_line_in_PRISM_case_note("    * F0107 sent to NCP")
	If CP_welcome_ltr_check = checked then call write_new_line_in_PRISM_case_note("    * Case Opening - Welcome letter to CP")
	If CP_new_order_summary_check = checked then call write_new_line_in_PRISM_case_note("    * New Order Summary to CP")
	If CP_PIN_Notice_check = checked then call write_new_line_in_PRISM_case_note("    * F0999 - PIN Notice to CP")
	If CP_health_ins_verif_check = checked then call write_new_line_in_PRISM_case_note("    * F0924 - Health Insurance Verification to CP")
	If child_care_verif_check = checked then call write_new_line_in_PRISM_case_note("    * Child Care Verification to CP")
	If CP_Stmt_of_Arrears_check = checked then call write_new_line_in_PRISM_case_note("    * Statement of Arrears Letter to CP")
	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note("* The following worklists created:")
	If t_10_day_tickler_check = checked then call write_new_line_in_PRISM_case_note("    * 10 day tickler to call NCP")
	If t_30_day_to_load_arrears_check = checked then call write_new_line_in_PRISM_case_note("    * 30 day tickler to load arrears")
	If t_30_day_case_review_check = checked then call write_new_line_in_PRISM_case_note("    * 30 day case review")	
	If t_60_day_case_review_check = checked then call write_new_line_in_PRISM_case_note("    * 60 day case review")	
	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note(add_caad_txt)
	call write_new_line_in_PRISM_case_note(worker_signature)
'	transmit

script_end_procedure("")
