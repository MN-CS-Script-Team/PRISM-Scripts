'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - INTAKE.vbs"
start_time = timer

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

'DIALOGS====================================================================================================================

BeginDialog intake_initial_dialog, 0, 0, 206, 85, "Intake: Initial Dialog"
  Text 5, 5, 50, 10, "Intake Script"
  Text 15, 20, 45, 10, "Case Number"
  EditBox 80, 15, 115, 15, PRISM_case_number
  Text 15, 40, 70, 10, "Type of Intake Action"
  DropListBox 100, 40, 95, 15, "Establishment"+chr(9)+"Enforcement"+chr(9)+"Motion to Set"+chr(9)+"Paternity", type_intake_drpdwn
  ButtonGroup ButtonPressed
    OkButton 95, 65, 50, 15
    CancelButton 150, 65, 50, 15
EndDialog

BeginDialog intake_enforcement_dialog, 0, 0, 391, 335, "Intake: Enforcement Dialog"
  Text 5, 5, 95, 10, "Intake: Enforcement"
  GroupBox 5, 20, 185, 90, "Documents Sending to CP:"
  CheckBox 15, 35, 145, 10, 	"Case Opening - Welcome Letter (*.docx)", CP_welcome_ltr_checkbox
  CheckBox 15, 50, 130, 10, 	"Child Care Verification (*.docx)", child_care_verif_checkbox
  CheckBox 15, 65, 140, 10, 	"Court Order Summary Letter (*.docx)", CP_new_order_summary_checkbox
  CheckBox 15, 80, 155, 10, 	"Health Insurance Verification (F0924)", CP_healthverif_checkbox
  CheckBox 15, 95, 100, 10, 	"Pin Notice (F0999)", CP_pinnotice_checkbox
  GroupBox 200, 20, 185, 90, "Documents Sending to NCP:"
  CheckBox 205, 35, 110, 10, 	"Arrears Amount Letter (*.docx)", NCP_arrearsletter_checkbox
  CheckBox 205, 50, 150, 10, 	"Case Opening - Introduction Letter (*.docx)", NCP_welcomeltr_checkbox
  CheckBox 205, 65, 130, 10, 	"Court Order Summary Letter (*.docx)", NCP_courtordersummary_checkbox
  CheckBox 205, 80, 155, 10, 	"Health Insurance Verfication (F0924)", NCP_healthverif_checkbox
  CheckBox 205, 95, 100, 10, 	"Pin Notice (F0999)", NCP_pinnotice_checkbox
  GroupBox 200, 115, 185, 110, "Send Liability Notice to NCP:"
  Text 210, 125, 50, 10, "NPA Cases"
  CheckBox 215, 135, 140, 10, 	"Authorization to Collect Support (F0100)", authtocollect_F0100_checkbox
  CheckBox 215, 150, 170, 10, 	"Notice of Child Support/Spousal Liability (F0108)", notice_csandspousal_dord_F0108_checkbox
  Text 210, 165, 50, 10, "PA Cases"
  CheckBox 215, 175, 165, 10, 	"Notice of Parental Liability for Support (F0109)", notification_of_parental_liability_dord_F0109_checkbox
  Text 210, 195, 50, 10, "MA Only Cases"
  CheckBox 215, 205, 165, 10, 	"Notice of Medical Support Liability (F0107)", notification_of_medical_liability_dord_F0107_checkbox
  GroupBox 5, 115, 185, 125, "CAWD Notes to Add:"
  Text 10, 130, 50, 10, "Worklist Text:"
  EditBox 65, 125, 110, 15, worklist_text_01
  Text 10, 145, 80, 10, "Calendar days until due:"
  EditBox 95, 140, 20, 15, calendar_days_until_due_01
  Text 10, 170, 50, 10, "Worklist Text:"
  EditBox 65, 165, 110, 15, worklist_text_02
  Text 10, 185, 80, 10, "Calendar days until due:"
  EditBox 95, 180, 20, 15, calendar_days_until_due_02
  Text 10, 210, 50, 10, "Worklist Text:"
  EditBox 65, 205, 110, 15, worklist_text_03
  Text 10, 225, 80, 10, "Calendar days until due:"
  EditBox 95, 220, 20, 15, calendar_days_until_due_03
  Text 205, 230, 75, 10, "File location on CAST:"
  EditBox 285, 225, 100, 15, file_location
  Text 205, 250, 180, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  EditBox 205, 275, 180, 15, add_text
  Text 205, 300, 70, 10, "Sign your CAAD Note:"
  EditBox 280, 295, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 280, 315, 50, 15
    CancelButton 335, 315, 50, 15
EndDialog

BeginDialog intake_establishment_dialog, 0, 0, 391, 330, "Intake: Establishment Dialog"
  Text 10, 5, 105, 10, "Intake: Establishment"
  GroupBox 5, 20, 185, 165, "Documents Sending to CP:"
  CheckBox 15, 30, 110, 15, 	"Child Care Expenses (*.docx)", daycare_checkbox
  CheckBox 15, 45, 80, 15, 		"Cover Letter (*.docx)", cover_letter_cp_checkbox
  CheckBox 15, 60, 120, 15, 	"Employment Verification (F0405)", emp_verif_cp_checkbox
  CheckBox 15, 75, 100, 15, 	"Financial Statement (F0021)", financial_stmt_cp_checkbox
  CheckBox 15, 90, 110, 15, 	"Medical Opinion Form (*.docx)", med_opinion_cp_checkbox
  CheckBox 15, 105, 120, 15, 	"Parenting Time Calendar (*.docx)", calendar_cp_checkbox
  CheckBox 15, 120, 100, 15, 	"Past Support Form (*.docx)", Past_support_cp_checkbox
  CheckBox 15, 135, 105, 15, 	"Statement of Rights (F0022)", stmt_right_cp_checkbox
  CheckBox 15, 150, 125, 15, 	"Waiver of Personal Service (F5000)", Waiver_cp_checkbox
  CheckBox 15, 165, 105, 15, 	"Your Privacy Rights (F0018)", priv_rights_cp_checkbox
  GroupBox 200, 20, 185, 195, "Documents Sending to NCP:"
  CheckBox 210, 30, 145, 15, 	"Authorization to Collect Support (F0100)", auth_collect_ncp_checkbox
  CheckBox 210, 45, 80, 15, 	"Cover Letter (*.docx)", cover_letter_ncp_checkbox
  CheckBox 210, 60, 120, 15, 	"Employment Verification (F0405)", emp_verif_ncp_checkbox
  CheckBox 210, 75, 105, 15, 	"Financial Statement (F0021)", financial_stmt_ncp_checkbox
  CheckBox 210, 90, 110, 15, 	"Medical Opinion Form (*.docx)", med_opinion_ncp_checkbox
  CheckBox 210, 105, 120, 15, 	"Parenting Time Calendar (*.docx)", calendar_ncp_checkbox
  CheckBox 210, 120, 100, 15, 	"Past Support Form (*.docx)", past_support_ncp_checkbox
  CheckBox 210, 135, 150, 15, 	"Notice of Medical Support Liability (F0107)", par_med_liab_ncp_checkbox
  CheckBox 210, 150, 160, 15, 	"Notice of Parental Liability for Support (F0109)", par_lia_ncp_checkbox
  CheckBox 210, 165, 105, 15, 	"Statement of Rights (F0022)", stmt_right_ncp_checkbox
  CheckBox 210, 180, 130, 15, 	"Waiver of Personal Service (F5000)", waiver_ncp_checkbox
  CheckBox 210, 195, 105, 15, 	"Your Privacy Rights (F0018)", priv_rights_ncp_checkbox
  GroupBox 5, 190, 185, 125, "CAWD Notes to Add:"
  Text 10, 205, 50, 10, "Worklist Text:"
  EditBox 65, 200, 120, 15, worklist_text_01
  Text 10, 220, 80, 10, "Calendar days until due:"
  EditBox 95, 215, 20, 15, calendar_days_until_due_01
  Text 10, 245, 50, 10, "Worklist Text:"
  EditBox 65, 240, 120, 15, worklist_text_02
  Text 10, 260, 80, 10, "Calendar days until due:"
  EditBox 95, 255, 20, 15, calendar_days_until_due_02
  Text 10, 285, 50, 10, "Worklist Text:"
  EditBox 65, 280, 120, 15, worklist_text_03
  Text 10, 300, 80, 10, "Calendar days until due:"
  EditBox 95, 295, 20, 15, calendar_days_until_due_03
  Text 205, 225, 75, 10, "File location on CAST:"
  EditBox 285, 220, 100, 15, file_location
  Text 205, 245, 180, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  EditBox 205, 270, 180, 15, add_text
  Text 205, 295, 70, 10, "Sign your CAAD Note:"
  EditBox 280, 290, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 280, 310, 50, 15
    CancelButton 335, 310, 50, 15
EndDialog

BeginDialog intake_motion_to_set_dialog, 0, 0, 391, 285, "Intake: Motion to Set Dialog"
  Text 5, 5, 75, 10, "Intake: Motion to Set"
  GroupBox 5, 20, 185, 90, "Documents Sending to CP:"
  CheckBox 10, 30, 145, 10, 	"Case Opening - Welcome Letter (*.docx)", CP_coverletter_checkbox
  CheckBox 10, 45, 115, 10, 	"Finacial Statement (F0021)", CP_Finacial_Statement_checkbox
  CheckBox 10, 60, 130, 10, 	"Child Care Verification (*.docx)", child_care_verif_checkbox
  CheckBox 10, 75, 140, 10, 	"Medical Opinion Form (*.docx)", CP_Medical_opinion_checkbox
  CheckBox 10, 90, 120, 10, 	"Employment Verification (F0405)", CP_Employment_Verification_checkbox
  GroupBox 200, 20, 185, 55, "Documents Sending to NCP:"
  CheckBox 205, 30, 120, 10, 	"Finacial Statement (F0021)", NCP_finacial_statement_checkbox
  CheckBox 205, 60, 115, 10, 	"Employment Verification (F0405)", NCP_employment_verification_checkbox
  CheckBox 205, 45, 130, 10, 	"Medical Opinion Form (*.docx)", NCP_medical_opinion_checkbox
  GroupBox 200, 80, 185, 90, "Send Liability Notice to NCP:"
  Text 210, 90, 40, 10, "NPA Cases"
  CheckBox 215, 100, 140, 10, 	"Authorization to Collect Support (F0100)", authorization_to_collect_support_F0100_checkbox
  Text 210, 115, 95, 10, "MFIP, DWP, or CCA cases"
  CheckBox 215, 130, 160, 10, 	"Notice of Parental Liability for Support (F0109)", notification_of_parental_liability_dord_F0109_checkbox
  Text 210, 145, 50, 10, "MA only cases"
  CheckBox 215, 155, 165, 10, 	"Notice of Medical Support Liability (F0107)", notification_of_medical_liability_dord_F0107_checkbox
  GroupBox 5, 115, 185, 125, "CAWD Notes to Add:"
  Text 10, 130, 50, 10, "Worklist Text:"
  EditBox 65, 125, 120, 15, worklist_text_01
  Text 10, 145, 80, 10, "Calendar days until due:"
  EditBox 95, 140, 20, 15, calendar_days_until_due_01
  Text 10, 170, 50, 10, "Worklist Text:"
  EditBox 65, 165, 120, 15, worklist_text_02
  Text 10, 185, 80, 10, "Calendar days until due:"
  EditBox 95, 180, 20, 15, calendar_days_until_due_02
  Text 10, 210, 50, 10, "Worklist Text:"
  EditBox 65, 205, 120, 15, worklist_text_03
  Text 10, 225, 80, 10, "Calendar days until due:"
  EditBox 95, 220, 20, 15, calendar_days_until_due_03
  Text 205, 180, 75, 10, "File location on CAST:"
  EditBox 285, 175, 100, 15, file_location
  Text 205, 200, 180, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  EditBox 205, 225, 180, 15, add_text
  Text 205, 250, 70, 10, "Sign your CAAD Note:"
  EditBox 280, 245, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 280, 265, 50, 15
    CancelButton 335, 265, 50, 15
EndDialog

BeginDialog intake_paternity_dialog, 0, 0, 391, 310, "Intake: Paternity Dialog"
  Text 5, 5, 85, 10, "Paternity Case Initiation"
  GroupBox 5, 20, 185, 150, "Documents Sending to CP:"
  CheckBox 15, 35, 110, 10, 	"Child Care Expenses (*.docx)", cp_child_care_expense_checkbox
  CheckBox 15, 50, 80, 10, 		"Cover Letter (*.docx)", cp_cover_letter_checkbox
  CheckBox 15, 65, 100, 10, 	"Financial Statement (F0021)", financial_stmt_cp_checkbox
  CheckBox 15, 80, 110, 10, 	"Medical Opinion Form (*.docx)", cp_medical_opinion_checkbox
  CheckBox 15, 95, 150, 10, 	"Paternity Questionnaire Affidavit (*.docx)", paternity_questionnaire_affidavit_checkbox
  CheckBox 15, 110, 100, 10, 	"Past Support Form (*.docx)", cp_past_support_form_checkbox
  CheckBox 15, 125, 100, 10, 	"Statement of Rights (F0022)", cp_statement_of_rights_checkbox
  CheckBox 15, 140, 125, 10, 	"Waiver of Personal Service (F5000)", cp_waiver_of_personal_service_checkbox
  CheckBox 15, 155, 130, 10, 	"Your Privacy Rights (F0018)", cp_your_privacy_rights_checkbox
  GroupBox 200, 20, 185, 135, "Documents Sending to NCP:"
  CheckBox 210, 35, 140, 10, 	"Cover Letter (*.docx)", ncp_cover_letter_checkbox
  CheckBox 210, 50, 145, 10, 	"Financial Statement (F0021)", financial_stmt_ncp_checkbox
  CheckBox 210, 65, 110, 10, 	"Medical Opinion Form (*.docx)", ncp_medical_opinion_checkbox
  CheckBox 210, 80, 150, 10, 	"NCP Voluntary Paternity Notice (F0516)", ncp_voluntary_paternity_notice_checkbox
  CheckBox 210, 95, 120, 10, 	"Past Support Form (*.docx)", ncp_past_support_form_checkbox
  CheckBox 210, 110, 110, 10, 	"Statement of Rights (F0022)", ncp_statement_of_rights_checkbox
  CheckBox 210, 125, 125, 10, 	"Waiver of Personal Service (F5000)", ncp_waiver_of_personal_service_checkbox
  CheckBox 210, 140, 130, 10, 	"Your Privacy Rights (F0018)", ncp_your_privacy_rights_checkbox
  GroupBox 5, 180, 185, 125, "CAWD Notes to Add:"
  Text 10, 195, 50, 10, "Worklist Text:"
  EditBox 65, 190, 120, 15, worklist_text_01
  Text 10, 210, 80, 10, "Calendar days until due:"
  EditBox 95, 205, 20, 15, calendar_days_until_due_01
  Text 10, 235, 50, 10, "Worklist Text:"
  EditBox 65, 230, 120, 15, worklist_text_02
  Text 10, 250, 80, 10, "Calendar days until due:"
  EditBox 95, 245, 20, 15, calendar_days_until_due_02
  Text 10, 275, 50, 10, "Worklist Text:"
  EditBox 65, 270, 120, 15, worklist_text_03
  Text 10, 290, 80, 10, "Calendar days until due:"
  EditBox 95, 285, 20, 15, calendar_days_until_due_03
  Text 205, 170, 75, 10, "File location on CAST:"
  EditBox 285, 165, 100, 15, file_location
  Text 205, 190, 180, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  EditBox 205, 215, 180, 15, add_text
  Text 205, 275, 70, 10, "Sign your CAAD Note:"
  EditBox 280, 270, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 280, 290, 50, 15
    CancelButton 335, 290, 50, 15
EndDialog


'SHOW THE INITIAL DIALOG=================================
DO
	err_msg = ""

	Dialog intake_initial_dialog
	if ButtonPressed = 0 then StopScript

	call PRISM_case_number_validation (PRISM_case_number, is_correct)
	if is_correct = false then err_msg = err_msg & vbnewline & "Invalid PRISM Case Number"
	if err_msg <> "" then msgbox "***NOTICE***" & err_msg
LOOP until err_msg = ""
