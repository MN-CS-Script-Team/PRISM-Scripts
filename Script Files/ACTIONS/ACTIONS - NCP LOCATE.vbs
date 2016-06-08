
'Gathering Stats==============================================================================================================================
name_of_script = "ACTION-NCP LOCATE.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 90
STATS_denomination = "C"
'End of STATS Block===========================================================================================================================

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

'THE DIALOG-----------------------------------------------------------------------------------------------------------------

BeginDialog Locate_dialog, 0, 0, 216, 375, "Locate"
  CheckBox 10, 40, 50, 10, "NCQW", NCQW_checkbox
  CheckBox 10, 55, 50, 10, "FIAD", FIAD_checkbox
  CheckBox 10, 70, 50, 10, "NCLA", NCLA_checkbox
  CheckBox 10, 85, 50, 10, "LOID", LOID_checkbox
  CheckBox 10, 100, 50, 10, "NCUI/FCUI", NCUIFCUI_checkbox
  CheckBox 75, 40, 50, 10, "DOLR", DOLR_checkbox
  CheckBox 75, 55, 70, 10, "SSSD/SSTD", SSA_checkbox
  CheckBox 75, 70, 110, 10, "NEBR", new_hire_checkbox
  CheckBox 75, 85, 50, 10, "NCMR", NCMR_checkbox
  CheckBox 75, 100, 80, 10, "Other PRISM Cases", Other_cases_checkbox
  CheckBox 10, 135, 50, 10, "MAXIS", MAXIS_checkbox
  CheckBox 10, 150, 50, 10, "MMIS", MMIS_checkbox
  CheckBox 10, 165, 50, 10, "MEC2", MEC2_checkbox
  CheckBox 75, 135, 50, 10, "DOC", DOC_checkbox
  CheckBox 75, 150, 85, 10, "Odyssey and/or MNCIS", Courts_checkbox
  CheckBox 75, 165, 50, 10, "DVS", DVS_checkbox
  CheckBox 10, 205, 50, 10, "Called CP", Called_CP_checkbox
  CheckBox 75, 205, 50, 10, "Called NCP", Called_NCP_checkbox
  CheckBox 10, 220, 165, 10, "Generate F0460 CP Locate Questionnaire", CP_Locate_checkbox
  CheckBox 10, 235, 160, 10, "Generate F0465 NCP Locate Questionnaire", NCP_Locate_checkbox
  CheckBox 10, 265, 110, 10, "Request header credit report", header_report_checkbox
  CheckBox 10, 280, 200, 10, "Generate F0444 Notice of Credit Bureau Inquiry (full report)", full_report_checkbox
  EditBox 60, 300, 150, 15, new_info
  EditBox 80, 335, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 105, 355, 50, 15
    CancelButton 155, 355, 50, 15
  Text 5, 10, 230, 15, "CHOOSE INFORMATION  REVIEWED AND ACTIONS TAKEN:"
  GroupBox 0, 25, 210, 90, "PRISM Screens"
  GroupBox 0, 120, 210, 60, "Systems/Websites"
  GroupBox 0, 185, 210, 60, "Client Contact"
  GroupBox 0, 250, 210, 45, "Credit Bureau"
  Text 5, 305, 55, 10, "New Info Found:"
  Text 5, 340, 70, 10, "Initials for CAAD note:"
  CheckBox 5, 320, 90, 10, "Add LOCATE worklist for", worklist_checkbox
  EditBox 95, 315, 30, 15, worklist_days
  Text 125, 320, 20, 10, "days"
EndDialog



'Enter a case number dialog so that if the case number is not already selected - the worker can enter one
BeginDialog case_number_dialog, 0, 0, 146, 45, "Dialog"
  EditBox 60, 5, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 40, 25, 50, 15
    CancelButton 95, 25, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

'THE SCRIPT---------------------------------------------------------------------------------------------------------------------

EMConnect ""

'getting case number for later use in script
call navigate_to_PRISM_screen("CAST")
EMWaitReady 0, 0
EMReadScreen PRISM_case_number, 13, 4, 8

Do
	Dialog case_number_dialog
	If ButtonPressed = 0 then StopScript
	If PRISM_case_number = "" then msgbox "Please enter a valid case number."
LOOP UNTIL PRISM_case_number <> ""


'navigating through the locate sceens on PRISM
call navigate_to_PRISM_screen("NCQW")
	ncqw_message = msgbox ("Review NCQW for recent wage info", vbOkCancel, "Check NCQW")
	If ncqw_message = vbCancel then stopscript
call navigate_to_PRISM_screen("FIAD")
	fiad_message = msgbox ("Review FIAD for address", vbOkCancel, "Check FIAD")
	If fiad_message = vbCancel then stopscript
call navigate_to_PRISM_screen("NCLA")
	ncla_message = msgbox ("Review for unreviewed locate attempts", vbOkCancel, "Check NCLA")
	If ncla_message = vbCancel then stopscript
call navigate_to_PRISM_screen("LOID")
	loid_message = msgbox ("Review Locate Detail", vbOkCancel, "Check LOID")
	If loid_message = vbCancel then stopscript
call navigate_to_PRISM_screen("NCUI")
	ncui_message = msgbox ("Review Unemployment Claims", vbOkCancel, "Check NCUI")
	If ncui_message = vbCancel then stopscript
call navigate_to_PRISM_screen("FCUI")
	fcui_message = msgbox ("Review Unemployment for NCP's in other states", vbOkCancel, "Check FCUI")
	If fcui_message = vbCancel then stopscript
call navigate_to_PRISM_screen("DOLR")
	dolr_message = msgbox ("Review DOC Locate Info", vbOkCancel, "Check DOLR")
	If dolr_message = vbCancel then stopscript
call navigate_to_PRISM_screen("SSTD")
	sstd_message = msgbox ("Review for RSDI Benefits", vbOkCancel, "Check SSTD")
	If sstd_message = vbCancel then stopscript
call navigate_to_PRISM_screen("SSSD")
	sssd_message = msgbox ("Review for SSI Benefits", vbOkCancel, "Check SSSD")
	If sssd_message = vbCancel then stopscript
call navigate_to_PRISM_screen("NCDE")		
	EMReadScreen ncp_ssn, 11, 6, 7 			'pulling ncp ssn from ncde
	ncp_ssn=replace(ncp_ssn, "-", "")			'replacing - in ssn with no space
call navigate_to_PRISM_screen("NEBR")	
	EmWriteScreen ncp_ssn, 20, 7				'writing ssn without - in SSN blank
	transmit
	nebr_message = msgbox ("Review New Hires", vbOkCancel, "Check NEBR")
	If nebr_message = vbCancel then stopscript
call navigate_to_PRISM_screen("NCMR")
	ncmr_message = msgbox ("Request Locate if appropriate", vbOkCancel, "Check NCMR")
	If ncmr_message = vbCancel then stopscript
call navigate_to_PRISM_screen("NCCB")
	nccb_message = msgbox ("Review NCP's other cases as needed", vbOkCancel, "Check NCCB")
	If nccb_message = vbCancel then stopscript
'rentering case # in case worker looks at ncp's other cases
call navigate_to_PRISM_screen("CAST")
	EMWriteScreen "D", 3, 29
	EMWriteScreen PRISM_case_number, 4, 8 
	EMWriteScreen right(PRISM_case_number, 2), 4, 19
	transmit 
call navigate_to_PRISM_screen("NCDD")
	other_message = msgbox ("Check the following Systems/Websites as needed:"& vbCR & "   -MAXIS" & vbCR & "   -MMIS" & vbCR & "   -MEC2" & vbCR & "   -DOC Websites" & vbCR & "   -Odyssey and/or MNCIS" & vbCR & "   -DVS Website", vbOkCancel, "Check Other") 
	If other_message = vbCancel then stopscript
'checkng for last time Locates and CB headers were done on case. 
call navigate_to_PRISM_screen("CAAT")
	EMWriteScreen "D0001", 20, 29
	transmit
	EMReadScreen cp_activity, 5, 8, 22
	EMReadScreen cp_locate_date, 8, 8, 12
	IF cp_activity <> "D0001" THEN cp_locate_date = "never sent"
	EMWriteScreen "D1465", 20, 29
	transmit
	EMReadScreen ncp_activity, 5, 8, 22
	EMReadScreen ncp_locate_date, 8, 8, 12
	IF ncp_activity <> "D1465" THEN ncp_locate_date = "never sent"
	EMWriteScreen "L0161", 20, 29
	transmit
	EMReadScreen credit_header, 5, 8, 22
	EMReadScreen credit_header_date, 8, 8, 12
	IF credit_header <> "L0161" THEN credit_header_date = "never requested"
'checking if address are known/unknown
call navigate_to_PRISM_screen("CPDD")
	EMReadScreen cp_address, 1, 10, 46
	If cp_address = "Y" Then cp_address_locate = "known"
	If cp_address = "N" Then cp_address_locate = "unknown"
call navigate_to_PRISM_screen("NCDD")
	EMReadScreen ncp_address, 1, 10, 46
	If ncp_address = "Y" Then ncp_address_locate = "known"
	If ncp_address = "N" Then ncp_address_locate = "unknown"			
	
locate_message = msgbox ("Past locate actions taken:"& vbCR & vbCR & "CP Locate last sent:  " & cp_locate_date & vbCR & "NCP Locate last sent:  " & ncp_locate_date & vbCR & "Credit Bureau Header requested:  " & credit_header_date & vbCR & vbCR & "CP Address:  " & cp_address_locate & vbCR & "NCP Address:  " & ncp_address_locate , vbOkCancel, "Locate Status")



' default Prism Screen checkboxes to checked
NCQW_checkbox = Checked
FIAD_checkbox = Checked
NCLA_checkbox = Checked
LOID_checkbox = Checked
NCUIFCUI_checkbox = Checked
DOLR_checkbox = Checked	
SSA_checkbox = Checked
new_hire_checkbox = Checked
NCMR_checkbox = Checked


	
DO
	err_msg = ""
	Dialog Locate_Dialog
	If ButtonPressed = 0 then StopScript
	If worker_signature = "" THEN err_msg = err_msg & "Please sign your CAAD note."
	If err_msg <> "" THEN
		Msgbox "Please sign your CAAD note."
	END IF
LOOP UNTIL err_msg = ""
	

'brings worker to DORD and creates DORD Doc for CP Locate
If CP_Locate_checkbox = checked THEN 
	call navigate_to_PRISM_screen("DORD") 
 	EMWriteScreen "C", 3, 29 
 	transmit 
	EMWriteScreen "A", 3, 29 
 	EMWriteScreen "F0460", 6, 36 
 	transmit 
END IF


'brings worker to DORD and creates DORD Doc for NCP Locate
If NCP_Locate_checkbox = checked THEN 
	call navigate_to_PRISM_screen("NCDE")		
	EMReadScreen ncp_MCI, 10, 4, 7 
	call navigate_to_PRISM_screen ("DORD") 
 	EMWriteScreen "C", 3, 29 
 	transmit 
	EMWriteScreen "A", 3, 29
	EmWriteScreen ncp_MCI, 4, 15 
	EmWriteScreen "__", 4, 26 
	EMWriteScreen "F0465", 6, 36 
 	transmit 
	call navigate_to_PRISM_screen("CAST")
	EMWriteScreen "D", 3, 29
	EMWriteScreen PRISM_case_number, 4, 8 
	EMWriteScreen right(PRISM_case_number, 2), 4, 19
	transmit 

END IF


'brings worker to DORD and creates DORD Doc for Full Credit Report
If full_report_checkbox = checked THEN 
	call navigate_to_PRISM_screen("DORD") 
 	EMWriteScreen "C", 3, 29 
 	transmit 
	EMWriteScreen "A", 3, 29 
 	EMWriteScreen "F0444", 6, 36 
 	transmit 

END IF


'adding CAAD note for header request
If header_report_checkbox = checked THEN
	call navigate_to_PRISM_screen("NCDE")		
	EMReadScreen ncp_MCI, 10, 4, 7 
	call navigate_to_PRISM_screen("CAAD") 
	PF5
 	EMWriteScreen "L0161", 4, 54 
	EmWriteScreen ncp_MCI, 10, 30  
 	transmit 




msgbox "CAAD L0161 Header Request CRB to Experian has been added to your case." & vbCR & vbCR & "Take any additional steps needed to complete request."

END IF


'adding script so screens checked show on one continuous line in CAAD
IF NCQW_checkbox = checked then screens_checked = screens_checked & "NCQW" & ", "
IF FIAD_checkbox = checked THEN screens_checked = screens_checked & "FIAD" & ", "
If NCLA_checkbox = checked THEN screens_checked = screens_checked & "NCLA" & ", "
If LOID_checkbox = Checked THEN screens_checked = screens_checked & "LOID" & ", "
If NCUIFCUI_checkbox = Checked THEN screens_checked = screens_checked & "NCUI/FCUI" & ", "
If DOLR_checkbox = Checked THEN screens_checked = screens_checked & "DOLR" & ", "
If SSA_checkbox = Checked THEN screens_checked = screens_checked & "SSSD/SSTD" & ", "
If new_hire_checkbox = Checked THEN screens_checked = screens_checked & "NEBR" & ", "
If NCMR_checkbox = Checked THEN screens_checked = screens_checked & "NCMR" & ", "
If Other_cases_checkbox = Checked THEN screens_checked = screens_checked & "& Other PRISM Cases." 

'adding script so systems/websites checked show on one continuous line in CAAD
If MAXIS_checkbox = Checked THEN systems_checked = systems_checked & "MAXIS" & ", "
If MMIS_checkbox = Checked THEN systems_checked = systems_checked & "MMIS" & ", "
If MEC2_checkbox = Checked THEN systems_checked = systems_checked & "MEC2" & ", "
If DOC_checkbox = Checked THEN systems_checked = systems_checked & "DOC Websites" & ", "
If Courts_checkbox = Checked THEN systems_checked = systems_checked & "Odyssey/MNCIS" & ", "
If DVS_checkbox = Checked THEN systems_checked = systems_checked & "& DVS Website."

'adding script so client contact checked show on one continuous line in CAAD
If Called_CP_checkbox = Checked THEN client_contact = client_contact & "Called CP" & ", "
If Called_NCP_checkbox = Checked THEN client_contact = client_contact & "Called NCP" & ", "
If CP_locate_checkbox = Checked THEN client_contact = client_contact & "Sending Locate Form to CP" & ", "
If NCP_locate_checkbox = Checked THEN client_contact = client_contact & "Sending Locate Form to NCP." 

'adding script so credit bureau actions checked show on one continuous line in CAAD
If header_report_checkbox = Checked THEN credit_bureau = credit_bureau & "Requesting Credit Bureau Header" & ", "
If full_report_checkbox = Checked THEN credit_bureau = credit_bureau & "Sending Notice of Credit Bureau Inquiry To NCP." 

'adding CAWT note
If worklist_checkbox = Checked THEN 
Call navigate_to_PRISM_screen ("CAWT")
PF5
EMwritescreen "free", 4, 37
EMSetCursor 10, 4
CALL write_variable_in_CAAD ("Complete Locate Review")
EMWritescreen worklist_days, 17, 52
transmit 

End if 
'adding CAAD note
call navigate_to_PRISM_screen("CAAD") 
	PF5
 	EMWriteScreen "FREE", 4, 54 
	EMSetCursor 16, 4	
	CALL write_variable_in_CAAD ("Locate Review")
	CALL write_bullet_and_variable_in_CAAD ("PRISM screens Reviewed", screens_checked)
	CALL write_bullet_and_variable_in_CAAD ("Systems/Websites Reviewed", systems_checked)
	CALL write_bullet_and_variable_in_CAAD ("client contact", client_contact)
	CALL write_bullet_and_variable_in_CAAD ("Credit Bureau", credit_bureau)
	CALL write_bullet_and_variable_in_CAAD ("New Info", new_info)
	CALL write_variable_in_CAAD (worker_signature)
Transmit
PF3

Script_end_procedure("")
