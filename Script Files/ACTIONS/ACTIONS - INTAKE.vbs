'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - INTAKE.vbs"
start_time = timer

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

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
checked = 1
unchecked = 0
CAWD_check = checked
CAAD_note_check = checked


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog CS_intake_dialog, 0, 0, 370, 344, "CS intake dialog"
  EditBox 60, 0, 90, 20, client_first_name
  EditBox 220, 0, 90, 20, client_last_name
  CheckBox 320, 0, 40, 20, "Caretaker", caretaker_checkbox
  EditBox 60, 30, 120, 20, street_line_1
  EditBox 60, 50, 120, 20, street_line_2
  EditBox 70, 70, 110, 20, city_state_zip
  EditBox 240, 30, 130, 20, NCP_name
  DropListBox 250, 50, 60, 20, "father"+chr(9)+"mother", NCP_gender
  EditBox 250, 70, 120, 20, childs_name
  Text 10, 120, 100, 10, "Paternity Cover Letter to CP:"
  CheckBox 20, 130, 60, 10, "Normal", paternity_cover_letter_normal_check
  CheckBox 20, 140, 80, 10, "Relative Caretaker", paternity_cover_letter_relative_caretaker_check
  CheckBox 20, 150, 100, 10, "Minor with GAL Attachment", paternity_cover_letter_minor_check
  CheckBox 10, 170, 150, 10, "Paternity Information Form Memo", paternity_information_form_memo_check
  CheckBox 10, 180, 150, 10, "Paternity Information Form", paternity_information_form_check
  CheckBox 10, 190, 150, 10, "Supplemental Paternity Information Form", supplemental_paternity_information_form_check
  CheckBox 10, 200, 160, 10, "Relative Caretaker Paternity Information Form", relative_caretaker_paternity_info_form_check
  CheckBox 10, 210, 150, 10, "CP Paternity Request Sheet", CP_paternity_request_sheet_check
  CheckBox 10, 220, 150, 10, "Issues-Paternity-to be Decided", issues_paternity_to_be_decided_check
  CheckBox 10, 230, 150, 10, "Parenting Time Schedules", parenting_time_schedules_check
  CheckBox 10, 240, 150, 10, "Financial Affidavit OCS", financial_affidavit_OCS_check
  CheckBox 10, 260, 130, 10, "Establishment Intake Ltr", Est_Ltr_checkbox
  CheckBox 200, 110, 100, 10, "F0018 - Your Privacy Rights", F0018_checkbox
  CheckBox 200, 120, 160, 10, "F0021 - Financial Statement", F0021_checkbox
  CheckBox 200, 130, 140, 10, "F0022 - Important Statement of Rights", F0022_check
  CheckBox 200, 140, 150, 10, "F0100 - Authorization to Collect Support", F0100_check
  CheckBox 200, 150, 170, 10, "F0109 - Notification of Parental Liability for Support", F0109_checkbox
  Text 200, 40, 40, 10, "NCP name:"
  EditBox 260, 180, 90, 20, worker_name
  EditBox 250, 200, 90, 20, worker_phone
  EditBox 280, 220, 90, 20, worker_signature
  CheckBox 200, 240, 160, 10, "Check here to have script send a CAAD note.", CAAD_note_check
  CheckBox 200, 250, 160, 10, "Check here to have script send a CAWD.", CAWD_check
  ButtonGroup ButtonPressed
    OkButton 220, 260, 50, 20
    CancelButton 280, 260, 50, 20
  CheckBox 200, 160, 150, 10, "F5000 - Waiver of Personal Service and Ltr", F5000_checkbox
  Text 10, 80, 70, 10, "City, state and zip:"
  GroupBox 0, 100, 180, 240, "Word docs to print for client"
  GroupBox 190, 20, 180, 70, "Familial info:"
  Text 200, 80, 50, 10, "Child's name:"
  Text 10, 40, 50, 10, "Street (line 1):"
  Text 200, 60, 50, 10, "NCP gender:"
  Text 200, 200, 50, 10, "Worker phone:"
  Text 200, 180, 50, 10, "Worker name:"
  Text 160, 10, 60, 10, "Client last name:"
  Text 0, 10, 60, 10, "Client first name:"
  GroupBox 0, 20, 180, 70, "Address"
  GroupBox 180, 100, 190, 70, "DORD docs to print for client"
  Text 200, 220, 70, 10, "Sign your CAAD note:"
  Text 10, 60, 50, 10, "Street (line 2):"
  
EndDialog
'CUSTOM FUNCTIONS***************************************************************************************************************
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

'This is a custom function to send a DORD doc to a particular recipient.  The two parameters are strings for the 
'dord doc form number to be generated and the 3-digit recipient code
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

'Reading info from CAPS
row = 6									'Setting these for an EMSearch for the comma, which signifies the end of a first name. Needs to start looking on row 6.
col = 1									'The col is where the search string must end, need to set as a 1 to show change (it shows a 0 for not found, signfying an error)
EMSearch ", ", row, col							'Finds that comma
last_name_length = col - 12						'12 being the starting column of the CP name
EMReadScreen client_last_name, last_name_length, 6, 12	'Reads the last name based on the length found above
EMReadScreen client_first_name, 20, 6, col + 2			'20 is a nice long amount, and the col variable contains the comma, we need to start reading two columns after that variable
client_first_name = trim(client_first_name)			'Getting rid of the excess spaces
EMReadScreen CH_MCI, 10, 18, 5 						'Reading the child's MCI
EMReadScreen NC_MCI, 10, 8, 11 						'Reading NCP's MCI

'Getting worker info for case note
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
EMReadScreen worker_phone, 12, 8, 35
PF3

'Get first child's name
call navigate_to_PRISM_screen("CHDE")
EMWriteScreen CH_MCI, 4, 7
transmit
EMReadScreen CH_F, 12, 9, 34
EMReadScreen CH_M, 12, 9, 56
EMReadScreen CH_L, 17, 9, 8
EMReadScreen CH_S, 3, 9, 74
childs_name = fix_read_data(CH_F) & " " & fix_read_data(CH_M) & " " & fix_read_data(CH_L)	
If trim(CH_S) <> "" then childs_name = childs_Name & " " & ucase(fix_read_data(CH_S))

'Go back to CAPS for all the kids' info
call navigate_to_PRISM_screen("CAPS")
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit	
'Getting all child/DOB info
PRISM_row = 18
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

' Get NCP's name
call navigate_to_PRISM_screen("NCDE")
EMWriteScreen NC_MCI, 4, 7
EMReadScreen NCP_F, 12, 8, 34
EMReadScreen NCP_M, 12, 8, 56
EMReadScreen NCP_L, 17, 8, 8
EMReadScreen NCP_S, 3, 8, 74
NCP_name = fix_read_data(NCP_F) & " " & fix_read_data(NCP_M) & " " & fix_read_data(NCP_L)	
If trim(NCP_S) <> "" then NCP_name = NCP_Name & " " & ucase(fix_read_data(NCP_S))
call navigate_to_PRISM_screen("CAPS")
EMReadScreen NCP_PRISM_gender,  3, 18, 76 			'Reading the relationship (FAT if it's father)
If NCP_PRISM_gender = "MOT" then					'If the NCP is "MOT" than NCP should read mother, otherwise it's reversed. It's almost always reversed.
	NCP_gender = "mother"
Else
	NCP_gender = "father"
End if

'Lower case-ing CP's names
call fix_case(client_first_name, 1)	
call fix_case(client_last_name, 1)

'Cleaning up worker info
worker_name = trim(worker_name)
call fix_case(worker_name, 1)
worker_name = change_client_name_to_FML(worker_name)

'Navigating to CPDD to pull address info
call navigate_to_PRISM_screen("CPDD")
EMReadScreen street_line_1, 30, 15, 11
EMReadScreen street_line_2, 30, 16, 11
EMReadScreen city_state_zip, 49, 17, 11

'Cleaning up address info
street_line_1 = replace(street_line_1, "_", "")
call fix_case(street_line_1, 1)
street_line_2 = replace(street_line_2, "_", "")
call fix_case(street_line_2, 1)
city_state_zip = replace(replace(replace(city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(city_state_zip, 2)

'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
Do
	Dialog CS_intake_dialog
	If buttonpressed = 0 then stopscript
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

'Combining variables from dialog
If street_line_2 <> "" then							'Address
	street_address = street_line_1 & chr(13) & street_line_2
Else
	street_address = street_line_1
End if
CP_name = trim(client_first_name) & trim(client_last_name)		'CP name

'Creating the Word application object (if any of the Word options are selected), and making it visible 
If _
	Est_Ltr_checkbox = checked or _
	CP_paternity_request_sheet_check = checked or _
	financial_affidavit_OCS_check = checked or _
	paternity_cover_letter_normal_check = checked or _
	paternity_cover_letter_relative_caretaker_check = checked or _
	paternity_cover_letter_minor_check = checked or _
	paternity_information_form_memo_check = checked or _
	paternity_information_form_check = checked or _
	relative_caretaker_paternity_info_form_check = checked or _
	supplemental_paternity_information_form_check = checked then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End if


'Updating the CP paternity request document
If CP_paternity_request_sheet_check = checked then
	
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\CP Paternity Request Sheet.dotx")
	With objDoc
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_CP_name").Result = CP_name
		.FormFields("field_AF_name").Result = NCP_name
		.FormFields("field_case_number").Result = PRISM_case_number
	End With
End if

'Updating the Financial Affidavit OCS document
If financial_affidavit_OCS_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Financial Affidavit OCS.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_all_children").Result = CAPS_kids
		.FormFields("field_CP_name").Result = CP_name
		'Must also add one of these: Name____________________________	Date of Birth_______________ for each kid, and autofill them!!!<<<<<<<<<<<
	End With
End if

'Updating the Normal Paternity Cover Letter to CP document
If paternity_cover_letter_normal_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Cover letter to CP - Normal.dotx")
	With objDoc
		.FormFields("field_name").Result = CP_name
		.FormFields("field_street_address").Result = street_address
		.FormFields("field_city_state_zip").Result = city_state_zip
		.FormFields("field_NCP_gender").Result = NCP_gender
		.FormFields("field_NCP_gender_02").Result = NCP_gender
		.FormFields("field_NCP_gender_03").Result = NCP_gender
		.FormFields("field_NCP_gender_04").Result = NCP_gender
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
		.FormFields("field_phone").Result = worker_phone
	End With
End if

'Opening the Relative Caretaker Paternity Cover Letter to CP document (does not autofill any info)
If paternity_cover_letter_relative_caretaker_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Cover letter to CP - Relative Caretaker.dotx")
	With objDoc
		.FormFields("field_name").Result = CP_name
		.FormFields("field_street_address").Result = street_address
		.FormFields("field_city_state_zip").Result = city_state_zip
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
		.FormFields("field_phone").Result = worker_phone

	End With
End if

'Updating the Minor with GAL Paternity Cover Letter to CP document
If paternity_cover_letter_minor_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Cover letter to CP - Minor with GAL attachment.dotx")
	With objDoc
		.FormFields("field_name").Result = CP_name
		.FormFields("field_street_address").Result = street_address
		.FormFields("field_city_state_zip").Result = city_state_zip
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
		.FormFields("field_phone").Result = worker_phone
		.FormFields("field_name_02").Result = CP_name
		.FormFields("field_case_number_02").Result = PRISM_case_number
	End With
End if

'Updating the Establishment Intake Ltr
If Est_Ltr_checkbox = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Establishment Intake Letter.dotx")
	With objDoc
		.FormFields("CPName").Result = CP_name
		.FormFields("CP_address").Result = street_address
		.FormFields("CP_CSZ").Result = city_state_zip
		.FormFields("PRISM_No").Result = PRISM_case_number
		.FormFields("CPName_2").Result = CP_name
		.FormFields("Due_Date").Result = dateadd("d", date, 5)
		.FormFields("worker").Result = worker_name
	End With
End if

'Opening the Paternity Information Form Memo document (does not autofill any info)
If paternity_information_form_memo_check = checked then set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Information Form Memo.dotx")

'Opening the Paternity Information Form document 
If paternity_information_form_check = checked then 
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Information Form.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_childs_name_2").Result = childs_name
	End With
End if

'Opening the Supplemental Paternity Information Form document 
If supplemental_paternity_information_form_check = checked then 
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Supplemental Paternity Information Form.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_fathers_name").Result = NCP_name
	End With
End if
'Opening the Relative Caretaker Paternity Information Form document 
If relative_caretaker_paternity_info_form_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Supplemental Paternity Information Form.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_fathers_name").Result = NCP_name
	End With
End if

'If F0018 is indicated on the dialog then it navigates to DORD to send it.
If F0018_checkbox = checked then
	call send_dord_doc("NCP", "F0018")		
End if

'If F0100 is indicated on the dialog then it navigates to DORD to send it.
If F0100_check = checked then
	call send_dord_doc("NCP", "F0100")	
End if

'If F0022 is indicated on the dialog then it navigates to DORD to send it.
If F0022_check = checked then
	if caretaker_checkbox = unchecked then
	send_msg = MsgBox("Do you want to send the F0022 Important Statement of Rights to both parties? Click Yes for both, or click No to send it to CP only.", vbYesNo)
		If send_msg = vbYes Then
			call send_dord_doc("NCP", "F0022")
			call send_dord_doc("CPP", "F0022")
		else
			call send_dord_doc("CPP", "F0022")		
		End If
	else
		call send_dord_doc("NCP", "F0022")
	end if
End if

'If F5000 is indicated on the dialog then it navigates to DORD to send it.
If F5000_checkbox = checked then
	if caretaker_checkbox = unchecked then
		call navigate_to_PRISM_screen("DORD")
		EMWriteScreen "C", 3, 29
		transmit
		EMWriteScreen "A", 3, 29
		EMWriteScreen "F5000", 6, 36
		transmit
		Pf14
		EMWriteScreen "U", 20, 14
		transmit
		EMWriteScreen "S", 12, 5
		transmit
		EMWriteScreen "12", 16, 15
		transmit
		PF3
		EMWriteScreen "M", 3, 29
		transmit 
		PF3
	End If
End if
'If F0109 is indicated on the dialog then it navigates to DORD to send it.
If F0109_checkbox = checked then
	call send_dord_doc("NCP", "F0109")	
	Pf14	
	EMWriteScreen "U", 20, 14
	transmit
	EMWriteScreen "S", 7, 5
	transmit
	EMWriteScreen "x", 16, 15
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit
	PF3
End If
'If F0021 is indicated on the dialog then it navigates to DORD to send it.
If F0021_checkbox = checked then
	if caretaker_checkbox = unchecked then
		call send_dord_doc("NCP", "F0021")
		call send_dord_doc("CPP", "F0021")
	else
		call send_dord_doc("NCP", "F0021")		
	End if
End if
If CAAD_note_check = checked then

	'Going to CAAD, adding a new note
	call navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")


	'Setting the type
	EMWriteScreen "M2123", 4, 54

	'Setting cursor in write area and writing note details
	EMSetCursor 16, 4
	call write_new_line_in_PRISM_case_note("* Intake packet sent to CP with the following docs:")
	If child_only_MA_check = checked then call write_new_line_in_PRISM_case_note("    * Child Only MA Choice of Service letter")
	If child_only_MA_relative_caretaker_check = checked then call write_new_line_in_PRISM_case_note("    * Child Only MA - relative caretaker letter")
	If CP_and_child_MA_check = checked then call write_new_line_in_PRISM_case_note("    * CP and Child MA choice of service letter")
	If CP_paternity_request_sheet_check = checked then call write_new_line_in_PRISM_case_note("    * CP Paternity Request sheet")
	If financial_affidavit_OCS_check = checked then call write_new_line_in_PRISM_case_note("    * Financial Affidavit OCS")
	If issues_paternity_to_be_decided_check = checked then call write_new_line_in_PRISM_case_note("    * Issues-Paternity-to be Decided")
	If parenting_time_schedules_check = checked then call write_new_line_in_PRISM_case_note("    * Parenting Time Schedules")
	If paternity_cover_letter_normal_check = checked then call write_new_line_in_PRISM_case_note("    * Normal Paternity Cover Letter to CP")
	If paternity_cover_letter_relative_caretaker_check = checked then call write_new_line_in_PRISM_case_note("    * Relative Caretaker Paternity Cover Letter to CP")
	If paternity_cover_letter_minor_check = checked then call write_new_line_in_PRISM_case_note("    * Minor with GAL Attachment")
	If paternity_information_form_memo_check = checked then call write_new_line_in_PRISM_case_note("    * Paternity Information Form Memo")
	If paternity_information_form_check = checked then call write_new_line_in_PRISM_case_note("    * Paternity Information Form")
	If supplemental_paternity_information_form_check = checked then call write_new_line_in_PRISM_case_note("    * Supplemental Paternity Information Form")
	If Est_Ltr_checkbox = checked then call write_new_line_in_PRISM_case_note("    * Establishment Intake Letter")
	If F0018_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F0018")
	If F0021_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F0021")
	If F0022_check = checked then call write_new_line_in_PRISM_case_note("    * DORD F0022")
	If F0100_check = checked then call write_new_line_in_PRISM_case_note("    * DORD F0100")
	If F0109_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F0109")
	If F5000_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F5000")


	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note("* CP to return by " & dateadd("d", date, 5) & ".")
	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note(worker_signature)

	transmit
End if

If CAWD_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** Intake Docs due from CP", 10, 4
	EMWriteScreen dateadd("d", date, 7), 17, 21
	transmit
End if

script_end_procedure("")
