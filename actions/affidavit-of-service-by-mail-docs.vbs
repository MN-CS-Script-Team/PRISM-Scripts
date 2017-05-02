'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "affidavit-of-service-by-mail.vbs"
start_time = timer
'STATS_counter = 1
'STATS_manualtime =              'MANUAL TIME NEEDED TO COMPLETE THIS SCRIPT IS NEEDED
'STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------------


'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'Dialogs==============================================================
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog AffOfServDialog, 0, 0, 301, 400, "Affidavit of Service By Mail"
  Text 10, 5, 70, 10, "CP Safety Concerns:"
  EditBox 90, 5, 30, 15, CP_GCSC
  Text 125, 5, 75, 10, "NCP Safety Concerns:"
  EditBox 205, 5, 25, 15, NCP_GCSC
  Text 10, 20, 260, 10, "Attorney Information (CARE):                                                         Withdrawl Date:"
  Text 10, 30, 65, 10, "CP Attorney Name:"
  EditBox 75, 30, 135, 15, CP_Attorney
  EditBox 215, 30, 50, 15, CP_withdrawl_date
  Text 10, 50, 70, 10, "NCP Attorney Name:"
  EditBox 80, 50, 130, 15, NCP_Attorney
  EditBox 215, 50, 50, 15, NCP_withdrawl_date
  Text 45, 65, 205, 10, "Who do you want to send the Affidavit to? (check all that apply)"
  CheckBox 60, 75, 25, 10, "NCP", ncp_button
  CheckBox 90, 75, 20, 10, "CP", cp_button
  CheckBox 115, 75, 55, 10, "NCP Attorney", NCP_Attorney_button
  CheckBox 175, 75, 55, 10, "CP Attorney", CP_attorney_button
  Text 10, 90, 170, 10, "What documents were served? (Check all that apply)"
  CheckBox 10, 100, 110, 10, "Summons and Complaint", summons_and_complaint
  CheckBox 180, 100, 90, 10, "Notice of Registration", Notice_of_Registration
  CheckBox 10, 115, 130, 10, "Amended Summons and Complaint", Amended_Summons_and_Complaint
  CheckBox 180, 115, 100, 10, "Notice of Settlement Conf.", Notice_of_Settlement_Conference
  CheckBox 10, 130, 100, 10, "Findings/Conclusion/Order", Findings_Conclusion_Order
  CheckBox 180, 130, 95, 10, "Aff of Default and ID", Aff_of_Default_and_ID
  CheckBox 10, 145, 130, 10, "Amended Findings/Conclusion/Order", Amended_Findings_Conclusion_Order
  CheckBox 180, 145, 115, 10, "Case Financial Summary - CAFS", Case_Financial_Summary
  CheckBox 10, 160, 115, 10, "Motion", motion
  CheckBox 180, 160, 95, 10, "Case Information Sheet", Case_Information_Sheet
  CheckBox 10, 175, 125, 10, "Amended Motion", Amended_Motion
  CheckBox 180, 175, 115, 10, "Case Payment History", Case_Payment_History
  CheckBox 10, 190, 95, 10, "Supporting Affidavit", supporting_affidavit
  CheckBox 180, 190, 95, 10, "Confidential Info Form", Confidential_Info_Form
  CheckBox 10, 205, 110, 10, "Amended Supporting Affidavit", Amended_Supporting_Affidavit
  CheckBox 180, 205, 105, 10, "Sealed Financial Document", sealed_financial_doc
  CheckBox 10, 220, 95, 10, "Financial Statement", financial_statement
  CheckBox 180, 220, 110, 10, "Important Statement of Rights", Important_Statement_of_Rights
  CheckBox 10, 235, 90, 10, "DES Information", des_information
  CheckBox 180, 235, 105, 10, "Your Privacy Rights", Your_Privacy_Rights
  CheckBox 10, 250, 100, 10, "Genetic/Blood Test Order", Genetic_Blood_Test_Order
  CheckBox 180, 250, 95, 10, "Request for Hearing", Request_for_Hearing
  CheckBox 10, 265, 110, 10, "Genetic/Blood Test Results", Genetic_Blood_Test_results
  CheckBox 180, 265, 110, 10, "Notice of Judgment Renewal", Notice_of_Judgment_Renewal
  CheckBox 10, 280, 95, 10, "Notice of Intervention", Notice_of_Intervention
  CheckBox 180, 280, 100, 10, "Guidelines Worksheet", guidelines_worksheet
  CheckBox 10, 295, 80, 10, "Notice of Hearing", Notice_of_Hearing
  Text 10, 310, 65, 10, "Other (Line 1)"
  Text 100, 310, 65, 10, "Other (Line 2)"
  Text 190, 310, 40, 10, "Date Served"
  EditBox 10, 320, 80, 15, other_line_1
  EditBox 100, 320, 80, 15, other_line_2
  EditBox 190, 320, 85, 15, date_box
  Text 15, 340, 60, 10, "CP Confidential?"
  CheckBox 85, 340, 25, 10, "Yes", CP_confidential
  Text 135, 340, 65, 10, "NCP Confidential?"
  CheckBox 205, 340, 25, 10, "Yes", NCP_confidential_yes
  Text 15, 355, 85, 10, "Served By Certified Mail?"
  CheckBox 105, 355, 25, 10, "Yes", Certified_mail_yes
  Text 135, 355, 100, 10, "Per County Attorney Direction?"
  CheckBox 240, 355, 25, 10, "Yes", CAO_direction_yes
  Text 10, 370, 70, 10, "Worker's Signature:"
  EditBox 10, 380, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 380, 65, 15
    CancelButton 230, 380, 65, 15
EndDialog

'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		IF buttonpressed = 0 THEN stopscript
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	LOOP UNTIL case_number_valid = TRUE
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	IF PRISM_check <> "PRISM" THEN MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
LOOP UNTIL PRISM_check = "PRISM"

'Checking that the user is not in a timed out PRISM
CALL check_for_PRISM(FALSE)

Do
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

CALL navigate_to_PRISM_screen("CARE")					'grabbing the Attorney information form CARE Screen 
EMReadScreen CP_Attorney, 30, 10, 09
EMReadScreen CP_withdrawl_date, 10, 19, 19
EMReadScreen NCP_Attorney, 30, 10, 48
EMReadScreen NCP_withdrawl_date, 10, 19, 58

CALL navigate_to_PRISM_screen("GCSC")					'Grabbing safety concerns from GCSC screen
EMReadScreen CP_GCSC, 1, 12, 24
EMReadScreen NCP_GCSC, 1, 13, 24

Do
	err_msg = ""
	Dialog AffOfServDialog 'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF ncp_button = 0 AND cp_button = 0 AND NCP_Attorney_button = 0 AND CP_Attorney_button = 0 THEN err_msg = err_msg & vbNewline & "Please select the receipiant for your Affidavit."
		IF date_box = "" THEN err_msg = err_msg & vbNewline & "The date served must be completed." 
		IF summons_and_complaint = 0 AND Amended_Summons_and_Complaint = 0 AND Findings_Conclusion_Order = 0 AND Amended_Findings_Conclusion_Order = 0 AND motion = 0 AND Amended_Motion = 0 AND supporting_affidavit = 0 AND Amended_Supporting_Affidavit = 0 AND financial_statement = 0 AND des_information = 0 AND Genetic_Blood_Test_Order = 0 AND Genetic_Blood_Test_results = 0 AND Notice_of_Intervention = 0 AND Notice_of_Hearing = 0 AND Notice_of_Registration = 0 AND Notice_of_Settlement_Conference = 0 AND Aff_of_Default_and_ID = 0 AND Case_Financial_Summary = 0 AND Case_Payment_History = 0 AND Case_Information_Sheet = 0 AND Confidential_Info_Form = 0 AND sealed_financial_doc = 0 AND Important_Statement_of_Rights = 0 AND Your_Privacy_Rights = 0 AND Request_for_Hearing = 0 AND Notice_of_Judgment_Renewal = 0 AND guidelines_worksheet = 0 AND other_line_1 = "" AND other_line_2 = "" THEN err_msg = err_msg & vbNewline & "At least one document must be selected."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

CALL navigate_to_PRISM_screen("CAAS")					'Grabbing County Name
EMReadScreen County_served, 20, 9, 25

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

EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen worker_signature, 16, 15
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
If CP_confidential = 1 then
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "Y", 16, 15
	Transmit
Else
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "N", 16, 15
	Transmit
End IF
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen County_served, 16, 15
Transmit

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)

'pulling address into excel coverletter

EMReadScreen name_served, 29, 13, 40
EMReadScreen address_line1, 33, 14, 40
EMReadScreen address_line2, 33, 15, 40
EMReadScreen address_line3, 33, 16, 40
EMReadScreen address_line4, 33, 17, 40


'adding into excel worbook 
set objExcel = CreateObject("Excel.Application")
Call excel_open ("H:\Global Applications\Gateway Services\CSU\Window Envelop.xlsx", True, True, ObjExcel, objWorkbook)

objExcel.Cells(1, 1).Value = name_served
objExcel.Cells(2, 1).Value = address_line1
objExcel.Cells(3, 1).Value = address_line2
objExcel.Cells(4, 1).Value = address_line3
objExcel.Cells(5, 1).Value = address_line4
objExcel.Cells(23, 1).Value = PRISM_case_number
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
objSheet.PrintOut

objworkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

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

EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen worker_signature, 16, 15
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
If NCP_confidential_yes = 1 then
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "Y", 16, 15
	Transmit
Else
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "N", 16, 15
	Transmit
End IF
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen County_served, 16, 15
Transmit


EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)

'pulling address into excel coverletter

EMReadScreen name_served, 29, 13, 40
EMReadScreen address_line1, 33, 14, 40
EMReadScreen address_line2, 33, 15, 40
EMReadScreen address_line3, 33, 16, 40
EMReadScreen address_line4, 33, 17, 40


'adding into excel worbook 
set objExcel = CreateObject("Excel.Application")
Call excel_open ("H:\Global Applications\Gateway Services\CSU\Window Envelop.xlsx", True, True, ObjExcel, objWorkbook)

objExcel.Cells(1, 1).Value = name_served
objExcel.Cells(2, 1).Value = address_line1
objExcel.Cells(3, 1).Value = address_line2
objExcel.Cells(4, 1).Value = address_line3
objExcel.Cells(5, 1).Value = address_line4
objExcel.Cells(23, 1).Value = PRISM_case_number
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
objSheet.PrintOut
objworkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

End If


'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if CP attorney checked

If CP_attorney_button = checked then
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

EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen worker_signature, 16, 15
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
If CP_confidential_yes = 1 then
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "Y", 16, 15
	Transmit
Else
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "N", 16, 15
	Transmit
End IF
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen County_served, 16, 15
Transmit

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)

'pulling address into excel coverletter

EMReadScreen name_served, 29, 13, 40
EMReadScreen address_line1, 33, 14, 40
EMReadScreen address_line2, 33, 15, 40
EMReadScreen address_line3, 33, 16, 40
EMReadScreen address_line4, 33, 17, 40


'adding into excel worbook 
set objExcel = CreateObject("Excel.Application")
Call excel_open ("H:\Global Applications\Gateway Services\CSU\Window Envelop.xlsx", True, True, ObjExcel, objWorkbook)

objExcel.Cells(1, 1).Value = name_served
objExcel.Cells(2, 1).Value = address_line1
objExcel.Cells(3, 1).Value = address_line2
objExcel.Cells(4, 1).Value = address_line3
objExcel.Cells(5, 1).Value = address_line4
objExcel.Cells(23, 1).Value = PRISM_case_number
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
objSheet.PrintOut
objworkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

End If

'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if NCP attorney checked

If NCp_attorney_button = checked then
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

EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen worker_signature, 16, 15
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
If NCP_confidential_yes = 1 then
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "Y", 16, 15
	Transmit
Else
	EMWriteScreen "S", 8, 5
	Transmit
	EMWriteScreen "N", 16, 15
	Transmit
End IF
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen County_served, 16, 15
Transmit

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

Dialog LH_dialog  'name of dialog
IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)

'pulling address into excel coverletter

EMReadScreen name_served, 29, 13, 40
EMReadScreen address_line1, 33, 14, 40
EMReadScreen address_line2, 33, 15, 40
EMReadScreen address_line3, 33, 16, 40
EMReadScreen address_line4, 33, 17, 40


'adding into excel worbook 
set objExcel = CreateObject("Excel.Application")
Call excel_open ("H:\Global Applications\Gateway Services\CSU\Window Envelop.xlsx", True, True, ObjExcel, objWorkbook)

objExcel.Cells(1, 1).Value = name_served
objExcel.Cells(2, 1).Value = address_line1
objExcel.Cells(3, 1).Value = address_line2
objExcel.Cells(4, 1).Value = address_line3
objExcel.Cells(5, 1).Value = address_line4
objExcel.Cells(23, 1).Value = PRISM_case_number
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
objSheet.PrintOut
objworkbook.Saved = True
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

End If

CALL navigate_to_PRISM_screen("CAAD")
PF5
EMWriteScreen "FREE", 4, 54
EMSetCursor 16, 4
call write_variable_in_CAAD("Service Details:")
IF CAO_direction_yes = 1 THEN call write_variable_in_CAAD("Service complete based on county attorney direction.")
IF ncp_button = 1 then call write_variable_in_CAAD("* NCP Served")
IF cp_button = 1 then call write_variable_in_CAAD("* CP Served")
IF NCP_Attorney_button = 1 then call write_bullet_and_variable_in_CAAD("NCP Attorney Served", NCP_Attorney)
IF CP_Attorney_button = 1 then call write_bullet_and_variable_in_CAAD("CP Attorney Served ", CP_Attorney)
IF CP_confidential = 1 then call write_variable_in_CAAD("CP Served Confidential")
IF NCP_confidential = 1 then call write_variable_in_CAAD("NCP Served Confidential")
call write_variable_in_CAAD(Worker_Signature)
Transmit
PF3

script_end_procedure("")
