'GATHERING STATS=================================
name_of_script = "case-initiation-docs-received.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/18/2017", "Worker Signature should now populate on this script.", "Kelly Hiestand, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'DIMMING VARIABLES
DIM row, col, case_number_valid, intake_docs_recd_dialog, paternity_wkst_check, rec_of_parentage_check, prism_case_number, date_recd, app_supp_coll_services_check, app_fee_check, ref_supp_coll_app_check, good_cause_check, role_county_atty_check, aff_arrears_check, waiver_pers_service_check, birth_check, marriage_check, court_order_check, photo_check, insurance_check, ButtonPressed, other_recd


'THE DIALOG BOX-------------------------------------------------------------------------------------------------------------------

BeginDialog case_initiation_docs_recd_dialog, 0, 0, 341, 180, "Case Initiation Documents Received"
  EditBox 75, 5, 75, 15, prism_case_number
  EditBox 225, 5, 65, 15, date_recd
  CheckBox 15, 40, 165, 10, "Application for Support/Coll Services DHS-1958", app_supp_coll_services_check
  CheckBox 15, 55, 155, 10, "Referral to Support/Collections DHS-3163B", ref_supp_coll_app_check
  CheckBox 15, 70, 85, 10, "Role of County Attorney", role_county_atty_check
  CheckBox 15, 85, 100, 10, "Waiver of Personal Service", waiver_pers_service_check
  CheckBox 15, 100, 105, 10, "Paternity Worksheet/Affidavit", paternity_wkst_check
  CheckBox 15, 115, 100, 10, "Marriage License/Certificate", marriage_check
  EditBox 40, 130, 130, 15, other_recd
  CheckBox 195, 40, 130, 10, "Client Stmt of Good Cause DHS-2338", good_cause_check
  CheckBox 195, 55, 105, 10, "Notarized Affidavit of Arrears", aff_arrears_check
  CheckBox 195, 70, 50, 10, "Court Order", court_order_check
  CheckBox 195, 85, 125, 10, "Health/Dental Insurance Verification", insurance_check
  CheckBox 195, 100, 95, 10, "Recognition of Parentage", rec_of_parentage_check
  CheckBox 195, 115, 30, 10, "Photo", photo_check
  CheckBox 195, 130, 65, 10, "Birth Verification", birth_check
  EditBox 85, 155, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 155, 50, 15
    CancelButton 280, 155, 50, 15
  Text 5, 10, 70, 10, "Prism Case Number:"
  Text 185, 10, 40, 15, "Date Rec'd:"
  Text 15, 130, 20, 10, "Other:"
  GroupBox 5, 25, 325, 125, "Documents Rec'd:"
  Text 10, 160, 70, 10, "Sign your CAAD note:"
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
	Dialog case_initiation_docs_recd_dialog
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
EMWriteScreen "*Case Initiation Documents Received*", 16, 4     'Types "Case Initiation Documents Received" on first line of CAAD note
EMSetCursor 17, 4                                               'Sets the cursor on the next line
IF date_recd <> "" THEN CALL write_bullet_and_variable_in_CAAD("Date Rec'd", date_recd)  'Types in date received on the second lind of CAAD note
IF app_supp_coll_services_check = checked THEN CALL write_variable_in_CAAD("* Application for Support/Coll Services DHS-1958")   'If any of the buttons are checked they will caad note
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
