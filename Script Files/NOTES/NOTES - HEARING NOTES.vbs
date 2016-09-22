'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HEARING NOTES.vbs"
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

'Using custom functions to convert arrays from global variables into a list for the dialogs.
call convert_array_to_droplist_items(county_attorney_array, county_attorney_list)										'County attorneys
call convert_array_to_droplist_items(child_support_magistrates_array, child_support_magistrates_list)					'County magistrates
call convert_array_to_droplist_items(county_judge_array, county_judge_list)												'County judges

'DIALOGS==================================================================================================================================================
'This dialog has been modified to show a dynamic county_attorney_list and child_support_magistrates_list from Global Variables. As such, it cannot be directly edited using dialog editor, without re-adding the preceding variable.
BeginDialog hearing_notes_expro_dialog, 0, 0, 321, 220, "Date of the Hearing ExPRO"
  Text 5, 5, 80, 10, "Motion before the Court"
  ComboBox 85, 5, 165, 15, "Select one or type in other motion:"+chr(9)+"MES 256 Action"+chr(9)+"Motion to Set"+chr(9)+"Continuance"+chr(9)+"License Suspension Appeal"+chr(9)+"COLA motion"+chr(9)+"Modification/RAM"+chr(9)+"UFM - Register for Modificaion", motion_before_court
  Text 5, 25, 85, 10, "Child Support Magistrate"
  DropListBox 90, 25, 85, 15, "Select one:" +chr(9)+ child_support_magistrates_list, child_support_magistrate
  Text 180, 25, 55, 10, "County Attorney"
  DropListBox 235, 25, 85, 15, "Select one:" +chr(9)+ county_attorney_list, CAO_list
  CheckBox 5, 50, 50, 10, "NCP present", NCP_present_check
  Text 60, 50, 60, 10, "Represented by:"
  EditBox 115, 50, 85, 15, NCP_represented_by
  CheckBox 5, 65, 50, 10, "CP present", CP_present_check
  Text 60, 65, 55, 10, "Represented by:"
  EditBox 115, 65, 85, 15, CP_represented_by
  Text 5, 90, 70, 10, "Details of the hearing"
  EditBox 75, 90, 170, 15, details_of_the_hearing
  CheckBox 5, 110, 100, 10, "Driver's license addressed", DL_addressed_check
  Text 20, 125, 105, 10, "Details of drivers license status"
  EditBox 130, 125, 155, 15, dl_details
  Text 10, 145, 70, 10, "Review Hearing Date"
  EditBox 85, 145, 65, 15, review_hearing_date
  Text 150, 175, 60, 10, "Worker signature"
  EditBox 215, 175, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 200, 50, 15
    CancelButton 255, 200, 50, 15
EndDialog

'This dialog has been modified to show a dynamic county_attorney_list and county_judge_list from Global Variables. As such, it cannot be directly edited using dialog editor, without re-adding the preceding variable.
BeginDialog hearing_notes_judicial_dialog, 0, 0, 321, 260, "Date of the Hearing Judicial"
  Text 5, 5, 80, 10, "Motion before the Court"
  ComboBox 85, 5, 155, 15, "Select one or type in other motion:"+chr(9)+"Initial Contempt of Court"+chr(9)+"Contempt Review"+chr(9)+"Continued Contempt Motion"+chr(9)+"Paternity Action", motion_before_court
  Text 5, 25, 65, 10, "District Court Judge"
  DropListBox 75, 25, 85, 15, "Select one:" +chr(9)+ county_judge_list, district_court_judge
  Text 5, 45, 55, 10, "County Attorney"
  DropListBox 65, 45, 85, 15, "Select one:" +chr(9)+ county_attorney_list, CAO_list
  CheckBox 5, 70, 50, 10, "NCP present", NCP_present_check
  Text 60, 70, 60, 10, "Represented by:"
  EditBox 120, 70, 85, 15, NCP_represented_by
  CheckBox 5, 85, 50, 10, "CP present", CP_present_check
  Text 60, 85, 55, 10, "Represented by:"
  EditBox 120, 85, 85, 15, CP_represented_by
  Text 5, 105, 70, 10, "Details of the hearing"
  EditBox 80, 105, 170, 15, details_of_the_hearing
  CheckBox 10, 125, 100, 10, "Driver's license addressed", DL_addressed_check
  Text 20, 140, 105, 10, "Details of drivers license status"
  EditBox 130, 135, 155, 15, dl_details
  Text 10, 160, 70, 10, "Review Hearing Date"
  EditBox 85, 160, 65, 15, review_hearing_date
  Text 150, 195, 60, 10, "Worker signature"
  EditBox 215, 195, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 210, 235, 50, 15
    CancelButton 265, 235, 50, 15
EndDialog

'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 135, "Case number dialog"
  EditBox 65, 75, 105, 15, PRISM_case_number
  DropListBox 105, 95, 65, 15, "Select one..."+chr(9)+"ExPro"+chr(9)+"Judicial", type_of_case_to_note_about
  ButtonGroup ButtonPressed
    OkButton 65, 115, 50, 15
    CancelButton 120, 115, 50, 15
  GroupBox 5, 5, 165, 65, "Info"
  Text 10, 15, 155, 50, "This script enters notes on CAAD using the M3909 (date of the hearing) code. Different pieces of information will be necessary for either an expedited process or judicial hearing action. Please select the type of hearing and press OK to continue."
  Text 10, 80, 50, 10, "Case number:"
  Text 10, 95, 90, 10, "Type of case to note about: "
EndDialog


'END DIALOGS=====================================================================================================================================================================================================


'Connecting to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)

'Case number display dialog
Do
	err_msg = ""																	'Blanking out the error message in case an error was previously detected, it needs to freshly re-evaluate during each run
	Dialog case_number_dialog														'Show the dialog
	If buttonpressed = 0 then stopscript											'Stop the script if cancel is pressed
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)			'Use a custom function to validate the case number entered

	'Handling error messages (incomplete entries) using an err_msg variable which gets added to if info is missing or incorrect
	If case_number_valid = False then err_msg = err_msg & "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''." & vbNewLine
	If type_of_case_to_note_about = "Select one..." then err_msg = err_msg & "You did not select a type of case to note about. Please select an option from the dropdown before continuing." & vbNewLine

	'Displaying the err_msg, will not finish the loop until err_msg is blank (meaning no errors detected)
	If err_msg <> "" then
		err_msg = "Please correct the following before pressing ''OK'' again:" & vbNewLine & vbNewLine & err_msg
		MsgBox err_msg
	End if

Loop until err_msg = ""		'If it's blank, we're good to move on to the next part because no errors were detected.


'If expro is selected, it will show the expro dialog. Otherwise it will show the judicial dialog.

If type_of_case_to_note_about = "ExPro" then

	'Displays dialog for hearing notes (expro) caad note and checks for information
	Do
		err_msg = ""							'Blanking out the error message in case an error was previously detected, it needs to freshly re-evaluate during each run
		Dialog hearing_notes_expro_dialog		'Show the dialog
		If buttonpressed = 0 then stopscript	'Stop the script if cancel is pressed

		'Handling error messages (incomplete entries) using an err_msg variable which gets added to if info is missing or incorrect
		If worker_signature = "" then err_msg = err_msg & "- You must sign your CAAD note." & vbNewLine																	'Makes sure worker enters in signature
		If motion_before_court = "" or motion_before_court = "Select one or type in other motion:" then err_msg = err_msg &  "- You must enter a motion." & vbNewLine 	'Makes sure worker selects motion type
		If CAO_list = "Select one:" then err_msg = err_msg &  "- You must select a County Attorney." & vbNewLine														'Makes sure worker selects county attorney
		If child_support_magistrate = "Select one:" then err_msg = err_msg & "- You must select a Child Support Magistrate."											'Makes sure worker selects child support magistrate

		'Displaying the err_msg, will not finish the loop until err_msg is blank (meaning no errors detected)
		If err_msg <> "" then
			err_msg = "Please correct the following before pressing ''OK'' again:" & vbNewLine & vbNewLine & err_msg
			MsgBox err_msg
		End if
	Loop until err_msg = ""		'If it's blank, we're good to move on to the next part because no errors were detected.

ElseIf type_of_case_to_note_about = "Judicial" then

	'Displays dialog for hearing notes (judicial) caad note and checks for information
	Do
		err_msg = ""							'Blanking out the error message in case an error was previously detected, it needs to freshly re-evaluate during each run
		Dialog hearing_notes_judicial_dialog	'Show the dialog
		If buttonpressed = 0 then stopscript	'Stop the script if cancel is pressed

		'Handling error messages (incomplete entries) using an err_msg variable which gets added to if info is missing or incorrect
		If worker_signature = "" then err_msg = err_msg & "- You must sign your CAAD note." & vbNewLine																	'Makes sure worker enters in signature
		If motion_before_court = "" or motion_before_court = "Select one or type in other motion:" then err_msg = err_msg &  "- You must enter a motion." & vbNewLine 	'Makes sure worker selects motion type
		If CAO_list = "Select one:" then err_msg = err_msg &  "- You must select a County Attorney." & vbNewLine														'Makes sure worker selects county attorney
		If district_court_judge = "Select one:" then err_msg = err_msg & "- You must select a District Court Judge."											'Makes sure worker selects child support magistrate

		'Displaying the err_msg, will not finish the loop until err_msg is blank (meaning no errors detected)
		If err_msg <> "" then
			err_msg = "Please correct the following before pressing ''OK'' again:" & vbNewLine & vbNewLine & err_msg
			MsgBox err_msg
		End if
	Loop until err_msg = ""		'If it's blank, we're good to move on to the next part because no errors were detected.

End if

'Checks for PRISM. If it's not found, the script ends here due to password out (or some other issue)
call check_for_PRISM(True)


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)


PF5					'Did this because you have to add a new note
EMWriteScreen "M3909", 4, 54  'adds correct caad code
EMSetCursor 16, 4			'Because the cursor does not default to this location

'Now we enter the CAAD note details from our dialog.

call write_bullet_and_variable_in_CAAD("Motion before the Court", motion_before_court)

'If...then logic used because ExPro uses magistrates while Judicial uses judges.
If type_of_case_to_note_about = "ExPro" then
	call write_bullet_and_variable_in_CAAD("Child Support Magistrate", child_support_magistrate)
ElseIf type_of_case_to_note_about = "Judicial" then
	call write_bullet_and_variable_in_CAAD("District Court Judge", district_court_judge)
End if

call write_bullet_and_variable_in_CAAD("County Attorney", CAO_list)
if NCP_present_check = 1 then
	call write_variable_in_CAAD("* NCP present")
	call write_bullet_and_variable_in_CAAD("Represented by", NCP_represented_by)
else
	call write_variable_in_CAAD ("* NCP not present")
end if
if CP_present_check = 1 then
	call write_variable_in_CAAD("* CP present")
	call write_bullet_and_variable_in_CAAD("Represented by", CP_represented_by)
else
	call write_variable_in_CAAD ("* CP not present")
end if
call write_bullet_and_variable_in_CAAD("Details of the Hearing", details_of_the_hearing)
if DL_addressed_check = 1 then
	call write_variable_in_CAAD("* Drivers license addressed")
	call write_bullet_and_variable_in_CAAD("Details of drivers license", dl_details)
end if
if review_hearing_date <> "" then
	call write_bullet_and_variable_in_CAAD("Review Hearing date", review_hearing_date)
end if
call write_variable_in_CAAD("---")
call write_variable_in_CAAD(worker_signature)

script_end_procedure("")
