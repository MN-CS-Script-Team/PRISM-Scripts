'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - AFFIDAVIT OF SERVICE BY MAIL.vbs"
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

Dim ncp_button, cp_button, ncp_attorney_button, cp_attorney_button, summons_and_complaint, Amended_Summons_and_Complaint, Findings_Conclusion_Order, des_information, Amended_Findings_Conclusion_Order, Amended_Motion, motion, supporting_affidavit, financial_statement, Amended_Supporting_Affidavit, Notice_of_Hearing, Genetic_Blood_Test_Order, Notice_of_Intervention, Genetic_Blood_Test_results, Notice_of_Registration, Notice_of_Settlement_Conference, Aff_of_Default_and_ID, Your_Privacy_Rights, Case_Financial_Summary, Case_Information_Sheet, Case_Payment_History, Confidential_Info_Form, Important_Statement_of_Rights, sealed_financial_doc, Request_for_Hearing, guidelines_worksheet, Notice_of_Judgment_Renewal, confidential_yes, confidential_no, date_box, certified_mail_yes, certified_mail_no, other_line_1, other_line_2
BeginDialog AffOfServDialog, 0, 0, 301, 380, "Affidavit of Service By Mail"
  ButtonGroup ButtonPressed
    OkButton 65, 345, 65, 15
    CancelButton 155, 345, 65, 15
  Text 60, 10, 50, 10, "Case Number:"
  Text 100, 45, 85, 10, " (Check all that apply)"
  Text 10, 60, 170, 10, "What documents were served? (Check all that apply)"
  Text 75, 25, 150, 10, "Who do you want to send the Affidavit to?"
  EditBox 110, 5, 110, 15, prism_case_number
  CheckBox 10, 75, 110, 10, "Summons and Complaint", summons_and_complaint
  CheckBox 10, 90, 130, 10, "Amended Summons and Complaint", Amended_Summons_and_Complaint
  CheckBox 10, 105, 100, 10, "Findings/Conclusion/Order", Findings_Conclusion_Order
  CheckBox 10, 120, 130, 10, "Amended Findings/Conclusion/Order", Amended_Findings_Conclusion_Order
  CheckBox 10, 135, 115, 10, "Motion", motion
  CheckBox 10, 150, 125, 10, "Amended Motion", Amended_Motion
  CheckBox 10, 165, 95, 10, "Supporting Affidavit", supporting_affidavit
  CheckBox 10, 180, 110, 10, "Amended Supporting Affidavit", Amended_Supporting_Affidavit
  CheckBox 10, 195, 95, 10, "Financial Statement", financial_statement
  CheckBox 10, 210, 90, 10, "DES Information", des_information
  CheckBox 10, 225, 100, 10, "Genetic/Blood Test Order", Genetic_Blood_Test_Order
  CheckBox 10, 240, 110, 10, "Genetic/Blood Test Results", Genetic_Blood_Test_results
  CheckBox 10, 255, 95, 10, "Notice of Intervention", Notice_of_Intervention
  CheckBox 10, 270, 80, 10, "Notice of Hearing", Notice_of_Hearing
  CheckBox 180, 75, 90, 10, "Notice of Registration", Notice_of_Registration
  CheckBox 180, 90, 100, 10, "Notice of Settlement Conf.", Notice_of_Settlement_Conference
  CheckBox 180, 105, 95, 10, "Aff of Default and ID", Aff_of_Default_and_ID
  CheckBox 180, 120, 115, 10, "Case Financial Summary - CAFS", Case_Financial_Summary
  CheckBox 180, 150, 115, 10, "Case Payment History", Case_Payment_History
  CheckBox 180, 135, 95, 10, "Case Information Sheet", Case_Information_Sheet
  CheckBox 180, 165, 95, 10, "Confidential Info Form", Confidential_Info_Form
  CheckBox 180, 180, 105, 10, "Sealed Financial Document", sealed_financial_doc
  CheckBox 180, 195, 110, 10, "Important Statement of Rights", Important_Statement_of_Rights
  CheckBox 180, 210, 105, 10, "Your Privacy Rights", Your_Privacy_Rights
  CheckBox 180, 225, 95, 10, "Request for Hearing", Request_for_Hearing
  CheckBox 180, 240, 110, 10, "Notice of Judgment Renewal", Notice_of_Judgment_Renewal
  CheckBox 180, 255, 100, 10, "Guidelines Worksheet", guidelines_worksheet
  CheckBox 60, 35, 25, 10, "NCP", ncp_button
  CheckBox 90, 35, 20, 10, "CP", cp_button
  CheckBox 115, 35, 55, 10, "NCP Attorney", NCP_Attorney_button
  CheckBox 175, 35, 55, 10, "CP Attorney", Cp_attorney_button
  OptionGroup RadioGroup1
    RadioButton 225, 325, 25, 10, "Yes", certified_mail_yes
    RadioButton 255, 325, 25, 10, "No", certified_mail_no
  OptionGroup RadioGroup2
    RadioButton 95, 325, 25, 10, "No", confidential_no
    RadioButton 65, 325, 25, 10, "Yes", confidential_yes
  EditBox 10, 305, 80, 15, other_line_1
  EditBox 100, 305, 80, 15, other_line_2
  Text 10, 295, 65, 10, "Other (Line 1)"
  Text 100, 295, 65, 10, "Other (Line 2)"
  Text 15, 325, 45, 10, "Confidential?"
  Text 190, 295, 40, 10, "Date Served"
  EditBox 190, 305, 85, 15, date_box
  Text 130, 325, 85, 10, "Served By Certified Mail?"
EndDialog


'Connects to Bluezone
EMConnect ""

'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if


'Starts dialog
					Dialog AffOfServDialog
     					IF ButtonPressed = 0 THEN StopScript
												
'goes to correct case
EMWriteScreen "CAST", 21,18
Transmit
EMWriteScreen "D", 3, 29
'Puts case number in from Dialog box
	EMWriteScreen Left (prism_case_number, 10), 4, 8
	EMWriteScreen Right (prism_case_number, 2), 4, 19
Transmit
'---------------------------------------------------------------------------------------------------------Creates DORD doc if CP checked
IF cp_button = checked then
'goes to DORD
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "CPP", 11, 51													
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit
EMSendKey (PF3)
End If
'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if NCP checked

If ncp_button = checked then
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "NCP", 11, 51													
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit
EMSendKey (PF3)

End If


'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if CP attorney checked

If Cp_attorney_button = checked then
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "CPA", 11, 51													
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit
EMSendKey (PF3)


End If

'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if NCP attorney checked

If Cp_attorney_button = checked then
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "NCA", 11, 51													
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then                                         						
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then                                         						
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit
EMSendKey (PF3)

End If

script_end_procedure("")


