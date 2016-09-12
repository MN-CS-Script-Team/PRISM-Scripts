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

BeginDialog initial_dialog, 0, 0, 206, 85, "Initial Dialog"
  Text 5, 5, 50, 10, "Intake Script"
  Text 15, 20, 45, 10, "Case Number"
  EditBox 80, 15, 115, 15, PRISM_case_number
  Text 15, 40, 70, 10, "Type of Intake Action"
  DropListBox 100, 40, 95, 15, "Establishment"+chr(9)+"Enforcement"+chr(9)+"Motion to Set"+chr(9)+"Paternity", type_intake_drpdwn
  ButtonGroup ButtonPressed
    OkButton 95, 65, 50, 15
    CancelButton 150, 65, 50, 15
EndDialog

BeginDialog Establisment_case_initiation_dialog, 0, 0, 381, 385, "Establishment Case Initiation"
  CheckBox 15, 40, 110, 15, "Child Care Expenses (*.docx)", daycare_checkbox
  CheckBox 15, 55, 80, 15, "Cover Letter (*.docx)", Cover_letter_cp_checkbox
  CheckBox 15, 70, 120, 15, "Employment Verification (F0405)", emp_verif_cp_checkbox
  CheckBox 15, 85, 100, 15, "Financial Statement (F0021)", financial_stmt_cp_checkbox
  CheckBox 15, 100, 110, 15, "Medical Opinion Form (*.docx)", med_opinion_cp_checkbox
  CheckBox 15, 115, 120, 15, "Parenting Time Calendar (*.docx)", calendar_cp_checkbox
  CheckBox 15, 130, 100, 15, "Past Support Form (*.docx)", Past_support_cp_checkbox
  CheckBox 15, 145, 105, 15, "Statement of Rights (F0022)", stmt_right_cp_checkbox
  CheckBox 15, 160, 125, 15, "Waiver of Personal Service (F5000)", Waiver_cp_checkbox
  CheckBox 15, 175, 105, 15, "Your Privacy Rights (F0018)", priv_rights_cp_checkbox
  CheckBox 205, 40, 145, 15, "Authorization to Collect Support (F0100)", auth_collect_ncp_checkbox
  CheckBox 205, 55, 80, 15, "Cover Letter (*.docx)", cover_letter_ncp_checkbox
  CheckBox 205, 70, 120, 15, "Employment Verification (F0405)", emp_verif_ncp_checkbox
  CheckBox 205, 85, 105, 15, "Financial Statement (F0021)", financial_stmt_ncp_checkbox
  CheckBox 205, 100, 110, 15, "Medical Opinion Form (*.docx)", med_opinion_ncp_checkbox
  CheckBox 205, 115, 120, 15, "Parenting Time Calendar (*.docx)", calendar_ncp_checkbox
  CheckBox 205, 130, 100, 15, "Past Support Form (*.docx)", past_support_ncp_checkbox
  CheckBox 205, 145, 150, 15, "Notice of Medical Support Liability (F0107)", par_med_liab_ncp_checkbox
  CheckBox 205, 160, 160, 15, "Notice of Parental Liability for Support (F0109)", par_lia_ncp_checkbox
  CheckBox 205, 175, 105, 15, "Statement of Rights (F0022)", stmt_right_ncp_checkbox
  CheckBox 205, 190, 130, 15, "Waiver of Personal Service (F5000)", waiver_ncp_checkbox
  CheckBox 205, 205, 105, 15, "Your Privacy Rights (F0018)", priv_rights_ncp_checkbox
  EditBox 70, 220, 110, 15, worklist_text_first
  EditBox 100, 235, 20, 15, cal_days_first
  EditBox 70, 260, 110, 15, worklist_text_second
  EditBox 100, 275, 20, 15, cal_days_second
  EditBox 70, 300, 110, 15, worklist_text_third
  EditBox 100, 315, 20, 15, cal_days_third
  EditBox 210, 245, 80, 15, file_location
  EditBox 195, 295, 140, 15, add_text
  EditBox 265, 325, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 230, 345, 50, 15
    CancelButton 295, 345, 50, 15
  Text 15, 320, 80, 10, "Calendar days until due:"
  GroupBox 5, 30, 175, 175, "Documents Sending to CP:"
  Text 15, 305, 50, 10, "Worklist Text:"
  Text 10, 10, 105, 10, "Establishment - Case Initiation"
  Text 15, 280, 80, 10, "Calendar days until due:"
  GroupBox 195, 30, 170, 205, "Documents Sending to NCP:"
  Text 15, 265, 50, 10, "Worklist Text:"
  GroupBox 200, 235, 105, 35, "File Location on CAST"
  Text 190, 330, 70, 10, "Sign your CAAD Note:"
  GroupBox 10, 210, 175, 135, "CAWD Notes to Add"
  Text 200, 275, 130, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  Text 15, 225, 50, 10, "Worklist Text:"
  Text 15, 240, 80, 10, "Calendar days until due:"
EndDialog





BeginDialog intake_enforcement_dialog, 0, 0, 401, 355, "Enforcement Intake Dialog"
  CheckBox 20, 35, 145, 10, "Case Opening - Welcome Letter (*.docx)", CP_welcome_ltr_checkbox
  CheckBox 20, 50, 130, 10, "Child Care Verification (*.docx)", child_care_verif_checkbox
  CheckBox 20, 65, 140, 10, "Court Order Summary Letter (*.docx)", CP_new_order_summary_checkbox
  CheckBox 20, 80, 155, 10, "Health Insurance Verification (F0924)", CP_healthverif_checkbox
  CheckBox 20, 95, 100, 10, "Pin Notice (F0999)", CP_pinnotice_checkbox
  CheckBox 215, 35, 110, 10, "Arrears Amount Letter (*.docx)", NCP_arrearsletter_checkbox
  CheckBox 215, 50, 150, 10, "Case Opening - Introduction Letter (*.docx)", NCP_welcomeltr_checkbox
  CheckBox 215, 65, 130, 10, "Court Order Summary Letter (*.docx)", NCP_courtordersummary_checkbox
  CheckBox 215, 80, 155, 10, "Health Insurance Verfication (F0924)", NCP_healthverif_checkbox
  CheckBox 215, 95, 100, 10, "Pin Notice (F0999)", NCP_pinnotice_checkbox
  CheckBox 225, 150, 140, 10, "Authorization to Collect Support (F0100)", authtocollect_F0100_checkbox
  CheckBox 225, 165, 170, 10, "Notice of Child Support/Spousal Liability (F0108)", notice_csandspousal_dord_F0108_checkbox
  CheckBox 225, 190, 165, 10, "Notification of Parental Liability Notice (F0109)", notification_of_parental_liability_dord_F0109_checkbox
  CheckBox 225, 220, 165, 10, "Notification of Medical Support Liability (F0107)", notification_of_medical_liability_dord_F0107_checkbox
  EditBox 35, 150, 135, 15, worklist_description_01
  EditBox 115, 170, 30, 15, calendar_days_01
  EditBox 35, 205, 135, 15, worklist_description_02
  EditBox 115, 225, 30, 15, calendar_days_02
  EditBox 35, 260, 135, 15, worklist_description_03
  GroupBox 210, 130, 185, 110, "Send Liability Notice to NCP:"
  Text 5, 5, 95, 10, "Enforcement Intake"
  GroupBox 210, 20, 185, 90, "Documents Sending to NCP:"
  Text 35, 285, 85, 10, "Calendar days until due"
  EditBox 115, 280, 30, 15, calendar_days_03
  EditBox 220, 250, 95, 15, filelocation_editbox
  EditBox 220, 290, 135, 15, caad_note_editbox
  EditBox 295, 310, 65, 15, signature_editbox
  ButtonGroup ButtonPressed
    OkButton 245, 335, 50, 15
    CancelButton 300, 335, 50, 15
  Text 35, 175, 80, 10, "Calendar days until due"
  Text 25, 195, 80, 10, "Worklist Description 2"
  Text 220, 210, 50, 10, "MA Only Cases"
  Text 35, 230, 85, 15, "Calendar days utnil due"
  Text 220, 140, 50, 10, "NPA Cases"
  Text 25, 250, 85, 10, "Worklist Description 3"
  Text 220, 180, 50, 10, "PA Cases"
  Text 220, 240, 75, 10, "File Location on CAST"
  GroupBox 10, 20, 180, 90, "Documents Sending to CP:"
  Text 220, 270, 135, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  GroupBox 15, 125, 175, 175, "CAWD Notes to Add:"
  Text 220, 315, 70, 10, "Sign your CAAD note:"
  Text 25, 140, 80, 15, "Worklist Description 1"
EndDialog

BeginDialog motion_to_set_intake_dialog, 0, 0, 381, 305, "Motion to Set Intake Dialog"
  Text 5, 0, 325, 10, "Motion to Set Case Initiation"
  GroupBox 10, 15, 170, 90, "Sending to CP"
  CheckBox 15, 25, 145, 10, "Case Opening - Welcome Letter (*.docx)", CP_coverletter_checkbox
  CheckBox 15, 40, 115, 10, "Finacial Statement (F0021)", CP_Finacial_Statement_checkbox
  CheckBox 15, 55, 130, 10, "Child Care Verification (*.docx)", child_care_verif_checkbox
  CheckBox 15, 70, 140, 10, "Medical Opinion Form (*.docx)", CP_Medical_opinion_checkbox
  CheckBox 15, 85, 120, 10, "Employment Verification (F0405)", CP_Employment_Verification_checkbox
  GroupBox 185, 15, 190, 140, "Sending to NCP"
  CheckBox 190, 25, 120, 10, "Finacial Statement (F0021)", NCP_finacial_statement_checkbox
  CheckBox 190, 55, 115, 10, "Employment Verification (F0405)", NCP_employment_verification_checkbox
  CheckBox 190, 40, 130, 10, "Medical Opinion Form (*.docx)", NCP_medical_opinion_checkbox
  Text 190, 75, 30, 10, "NPA"
  CheckBox 190, 85, 140, 10, "Authorization to Collect Support (F0100)", F0100_checkbox
  Text 190, 100, 95, 10, "MFIP, DWP, CCA"
  CheckBox 190, 115, 180, 10, "Notification of Parental Liability for Support (F0109)", F0109_checkbox
  Text 190, 130, 35, 10, "MA only"
  CheckBox 190, 140, 165, 10, "Notification for Medical Support Liability (F0107)", F0107_checkbox
  GroupBox 10, 110, 170, 190, "CAWD notes to add"
  Text 15, 125, 80, 10, "Worklist Description"
  EditBox 15, 140, 140, 15, Worklist_text_editbox_1
  Text 15, 160, 80, 10, "Calendar days until due"
  EditBox 100, 155, 30, 15, CAWT_number_days_editbox_1
  Text 15, 185, 75, 10, "Worklist Description"
  EditBox 15, 200, 140, 15, Worklist_text_editbox_2
  Text 15, 220, 80, 10, "Calendar days until due"
  EditBox 100, 215, 30, 15, CAWT_number_days_editbox_2
  Text 15, 245, 85, 10, "Worklist Description"
  EditBox 15, 260, 135, 15, Worklist_text_editbox_3
  Text 15, 280, 80, 10, "Calendar day until due"
  EditBox 100, 275, 30, 15, CAWT_number_days_editbox_3
  GroupBox 190, 160, 185, 35, "File Location on CAST"
  EditBox 195, 170, 40, 15, Location_CAST_textbox
  Text 190, 200, 165, 20, "Additional text to CAAD (Docs will automatically list in CAAD note)"
  EditBox 190, 225, 95, 15, CAAD_note_editbox
  Text 190, 255, 70, 10, "Sign your CAAD note:"
  EditBox 265, 250, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 255, 285, 50, 15
    CancelButton 310, 285, 50, 15
EndDialog

BeginDialog paternity_case_initiation_dialog, 0, 0, 356, 310, "Paternity Case Initiation"
  Text 5, 5, 85, 10, "Paternity Case Initiation"
  GroupBox 10, 15, 150, 150, "Letters to CP"
  CheckBox 10, 25, 100, 15, "Financial Statement (F0021)", financial_stmt_cp_checkbox
  CheckBox 10, 40, 80, 15, "Cover Letter (*.docx)", cp_cover_letter_checkbox
  CheckBox 10, 55, 130, 15, "Your Privacy Rights (F0018)", cp_your_privacy_rights_checkbox
  CheckBox 10, 70, 125, 15, "Waiver of Personal Service (F5000)", cp_waiver_of_personal_service_checkbox
  CheckBox 10, 85, 150, 15, "Paternity Questionnaire Affidavit (*.docx)", paternity_questionnaire_affidavit_checkbox
  CheckBox 10, 100, 100, 15, "Past Support Form (*.docx)", cp_past_support_form_checkbox
  CheckBox 10, 115, 110, 15, "Child Care Expenses (*.docx)", cp_child_care_expense_checkbox
  CheckBox 10, 130, 100, 15, "Statement of Rights (F0022)", cp_statement_of_rights_checkbox
  CheckBox 10, 145, 110, 15, "Medical Opinion Form (*.docx)", cp_medical_opinion_checkbox
  GroupBox 185, 15, 145, 130, "Letters to NCP"
  CheckBox 185, 25, 145, 15, "Financial Statement (F0021)", financial_stmt_ncp_checkbox
  CheckBox 185, 40, 140, 15, "Cover Letter (*.docx)", ncp_cover_letter_checkbox
  CheckBox 185, 55, 130, 15, "Your Privacy Rights (F0018)", ncp_your_privacy_rights_checkbox
  CheckBox 185, 70, 125, 15, "Waiver of Personal Service (F5000)", ncp_waiver_of_personal_service_checkbox
  CheckBox 185, 85, 150, 15, "NCP Voluntary Paternity Notice (F0516)", ncp_voluntary_paternity_notice_checkbox
  CheckBox 185, 100, 120, 15, "Past Support Form (*.docx)", ncp_past_support_form_checkbox
  CheckBox 185, 115, 110, 15, "Statement of Rights (F0022)", ncp_statement_of_rights_checkbox
  CheckBox 185, 130, 110, 15, "Medical Opinion Form (*.docx)", ncp_medical_opinion_checkbox
  GroupBox 5, 170, 140, 130, "CAWD notes to add"
  Text 10, 180, 75, 10, "Worklist Description"
  EditBox 10, 190, 130, 15, Edit2 t_30_day_cawd_txt
  Text 10, 205, 80, 10, "Calendar days until due"
  EditBox 90, 205, 30, 15, CAWT_number_days_editbox_1
  Text 10, 220, 75, 10, "Worklist Description"
  EditBox 10, 230, 130, 15, Edit9 worker_name
  Text 10, 245, 80, 10, "Calendar days until due"
  EditBox 90, 245, 30, 15, CAWT_number_days_editbox_2
  Text 10, 260, 75, 10, "Worklist Description"
  EditBox 10, 270, 130, 15, Edit11 worker_signature
  Text 10, 285, 80, 10, "Calendar days until due"
  EditBox 90, 285, 30, 15, CAWT_number_days_editbox_3
  GroupBox 180, 160, 80, 10, "File Location on CAST"
  EditBox 265, 155, 55, 15, Edit10 worker_phone
  Text 185, 175, 150, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD"
  EditBox 185, 195, 140, 15, Worklist_text_editbox_3
  Text 180, 225, 50, 10, "Worker name:"
  EditBox 230, 220, 90, 15, Worklist_text_editbox_1
  Text 180, 250, 50, 10, "Worker phone:"
  EditBox 230, 245, 90, 15, Worklist_text_editbox_2
  Text 185, 270, 70, 10, "Sign your CAAD note:"
  EditBox 255, 265, 40, 15, Location_CAST_textbox
  ButtonGroup ButtonPressed
    OkButton 245, 285, 50, 20
    CancelButton 300, 285, 50, 20
EndDialog


'SHOW THE INITIAL DIALOG=================================
DO
	err_msg = ""

	Dialog initial_dialog
	if ButtonPressed = 0 then StopScript

	call PRISM_case_number_validation (PRISM_case_number, is_correct)
	if is_correct = false then err_msg = err_msg & vbnewline & "Invalid PRISM Case Number"
	if err_msg <> "" then msgbox "***NOTICE***" & err_msg
LOOP until err_msg = ""



