'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "arrears-management-review.vbs"
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
call changelog_update ("02/22/2017", "Added Supervisor options.", "Jodi Martin, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog arrears_mgmt_dialog, 0, 0, 391, 385, "Arrears Mgmt Review"
  EditBox 170, 0, 80, 20, PRISM_case_number
  CheckBox 20, 50, 260, 10, "Check here if arrears mgmt is for CMS while NCP was also a recipient of MA", CMS_check
  EditBox 130, 60, 80, 15, cms_date_txt
  CheckBox 20, 80, 250, 10, "Check here if arrears mgmt is for charging while NCP rec'd cash assistance", Cash_PA_check
  EditBox 130, 90, 80, 15, cash_pa_date_txt
  CheckBox 20, 110, 250, 10, "Check here if arrears mgmt is for charging while NCP was incarcerated", incarcerated_check
  EditBox 130, 120, 80, 15, incar_date_txt
  CheckBox 20, 140, 190, 10, "Check here if arrears mgmt is for other circumstances ", other_check
  EditBox 260, 140, 80, 15, other_reason_txt
  EditBox 120, 150, 80, 15, other_date_txt
  EditBox 90, 180, 280, 15, details
  EditBox 110, 200, 260, 15, arrears_mgmt_amount
  CheckBox 30, 220, 330, 10, "Check here if your request for arrears mgmt includes suspension of PA interest charging", suspend_interest_check
  CheckBox 30, 230, 330, 10, "Check here if your request for arrears mgmt includes $1 PA forgiveness for every $1 payment", dollar_for_dollar_check
  Text 10, 10, 160, 10, "PRISM case number (XXXXXXXXXX-XX format):"
  EditBox 90, 295, 270, 35, add_info_txt
  EditBox 95, 350, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 360, 50, 15
    CancelButton 335, 360, 50, 15
  Text 100, 65, 30, 10, "Date(s):"
  DropListBox 95, 270, 265, 15, "Sent reqest to Supervisor"+chr(9)+"E9855 SUPERVISOR APPROVED ARREARS MANAGEMENT STRATEGY"+chr(9)+"E9856 SUPERVISOR APPROVED RECURRING ARREARS MGMT STRGY"+chr(9)+"E9857 SUPERVISOR DENIED ARREARS MANAGEMENT STRATEGY", decision_drop_down
  DropListBox 100, 30, 270, 20, "Supervisor decision "+chr(9)+"E9852 Reviewed for Arrears Mgmt - Approval Requested"+chr(9)+"E9851 Reviewed for Arrears Mgmt - No Action"+chr(9)+"E9853 Reviewed for Arrears Mgmt - More Information Needed"+chr(9)+"E9854 Arrears Management - Additional Information Not Returned"+chr(9)+"E9860 Arrears Management Recurring Strategy Ended"+chr(9)+"E9865 NPA CP Approved Arrears Management Strategy", CAAD_type
  Text 10, 30, 90, 10, "Please select CAAD note:"
  Text 125, 250, 105, 15, "*****Decision /  Response***** "
  ButtonGroup ButtonPressed
    PushButton 310, 0, 60, 15, "DHS Sir-Milo Info", DHS_sir_button
  Text 90, 150, 30, 10, "Date(s):"
  Text 15, 270, 75, 10, "Decision:"
  Text 10, 355, 70, 10, "Sign your case note: "
  Text 210, 140, 50, 10, "Reason detail:"
  Text 100, 95, 30, 10, "Date(s):"
  Text 10, 170, 80, 30, "Other details about this arrears mgmt review and CAAD note free text:"
  Text 100, 125, 30, 10, "Date(s):"
  Text 15, 290, 55, 15, "Additional info:"
  Text 10, 200, 100, 10, "Total amount of arrears mgmt:"
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
		error_msg = ""
		dialog arrears_mgmt_dialog
		IF buttonpressed = 0 THEN stopscript
		IF ButtonPressed = DHS_sir_button THEN
			CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/PRISM/User_docs/SIRMILO/Arrears_Management_Prevention_Policy/Pages/default.aspx")
		END IF
		IF other_check = checked and other_reason_txt = "" THEN
			error_msg = error_msg & vbCr & "Please enter reason detail for your arrears management request based on 'other circumstances'.  "
		END IF
		IF worker_signature = "" THEN
			error_msg = error_msg & vbCr & "Please sign your case note.  "
		END IF
		IF error_msg <> "" THEN
			Msgbox "Resolve to continue:" & vbCr & error_msg
		END IF


		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = FALSE THEN MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	LOOP UNTIL case_number_valid = TRUE and error_msg = "" and buttonpressed <> DHS_sir_button

	CALL navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	IF case_activity_detail <> "Case Activity Detail" THEN MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
LOOP UNTIL case_activity_detail = "Case Activity Detail"


IF CAAD_type = "E9852 Reviewed for Arrears Mgmt - Approval Requested" THEN CAAD_code = "E9852"
IF CAAD_type = "E9851 Reviewed for Arrears Mgmt - No Action" THEN CAAD_code = "E9851"
IF CAAD_type = "E9853 Reviewed for Arrears Mgmt - More Information Needed" THEN CAAD_code = "E9853"
IF CAAD_type = "E9854 Arrears Management - Additional Information Not Returned" THEN CAAD_code = "E9854"
IF CAAD_type = "E9860 Arrears Management Recurring Strategy Ended" THEN CAAD_code = "E9860"
IF CAAD_type = "E9865 NPA CP Approved Arrears Management Strategy" THEN CAAD_code = "E9865"
IF decision_drop_down = "E9855 SUPERVISOR APPROVED ARREARS MANAGEMENT STRATEGY" THEN CAAD_code = "E9855"
IF decision_drop_down = "E9856 SUPERVISOR APPROVED RECURRING ARREARS MGMT STRGY" THEN CAAD_code = "E9856"
IF decision_drop_down = "E9857 SUPERVISOR DENIED ARREARS MANAGEMENT STRATEGY" THEN CAAD_code = "E9857"



'Writing the case note
EMWriteScreen CAAD_code, 4, 54

EMSetCursor 16, 4 

'Because the PRISM case note functions require the cursor to start here
IF details <> "" THEN CALL write_bullet_and_variable_in_CAAD("Arrears Mgmt Review Details", details)
IF arrears_mgmt_amount <> "" THEN CALL write_bullet_and_variable_in_CAAD("Amount requested", arrears_mgmt_amount)
IF dollar_for_dollar_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt request includes $1 PA forgiveness for every $1 paid, if approved.")
IF suspend_interest_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt request includes suspension of PA interest charging, if approved.")

IF trim(cms_date_txt) <> "" THEN
	IF CMS_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt requested because CMS was charged while NCP was also a recipient of MA " & cms_date_txt & ".")
ELSE
	IF CMS_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt requested because CMS was charged while NCP was also a recipient of MA.")
END IF
IF trim(cash_pa_date_txt) <> "" THEN
	IF Cash_PA_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt requested because NCP received cash public assistance " & cash_pa_date_txt & ".")
ELSE
	IF Cash_PA_check = 1 THEN write_variable_in_CAAD("* Arrears Mgmt requested because NCP received cash public assistance.")
END IF
IF trim(incar_date_txt) <> "" THEN
	IF incarcerated_check = 1 THEN CALL write_variable_in_CAAD("* Arrears Mgmt requested because NCP was incarcerated " & incar_date_txt & ".")
ELSE
	IF incarcerated_check = 1 THEN CALL write_variable_in_CAAD("* Arrears Mgmt requested because NCP was incarcerated.")
END IF
IF trim(other_date_txt) <> "" THEN
	IF other_check = 1 THEN CALL write_variable_in_CAAD("* Arrears Mgmt requested: " & trim(other_reason_txt) &" " & other_date_txt & ".")
ELSE
	IF other_check = 1 THEN CALL write_variable_in_CAAD("* Arrears Mgmt requested: " & trim(other_reason_txt)  & ".")
END IF

IF add_info_txt <> "" THEN CALL write_variable_in_CAAD("* Additional information: " & trim(add_info_txt)  & ".")


	
CALL write_variable_in_CAAD("---")
CALL write_variable_in_CAAD(worker_signature)

script_end_procedure("")
