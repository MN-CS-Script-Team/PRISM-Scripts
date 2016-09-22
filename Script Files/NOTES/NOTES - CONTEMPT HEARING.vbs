'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CONTEMPT HEARING.vbs"
start_time = timer


'LOADING ROUTINE FUNCTIONS (FOR PRISM)--- UPDATED 9/8/16 to MASTER FUNCLIB--------------------------------------------------------------
IF IsEmpty(FuncLib_URL) = TRUE THEN 'Shouldn't load FuncLib if it already loaded once
    IF run_locally = FALSE or run_locally = "" THEN    'If the scripts are set to run locally, it skips this and uses an FSO below.
        IF use_master_branch = TRUE THEN               'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        Else                                            'Everyone else should use the release branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        End if
        SET req = CreateObject("Msxml2.XMLHttp.6.0")                'Creates an object to get a FuncLib_URL
        req.open "GET", FuncLib_URL, FALSE                          'Attempts to open the FuncLib_URL
        req.send                                                    'Sends request
        IF req.Status = 200 THEN                                    '200 means great success
            Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
            Execute req.responseText                                'Executes the script code
        ELSE                                                        'Error message
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
'END FUNCTIONS LIBRARY BLOCK=======

'Using custom functions to convert arrays from global variables into a list for the dialogs.
call convert_array_to_droplist_items(county_attorney_array, county_attorney_list)										'County attorneys
call convert_array_to_droplist_items(child_support_magistrates_array, child_support_magistrates_list)					'County magistrates
call convert_array_to_droplist_items(county_judge_array, county_judge_list)												'County judges



BeginDialog Contempt_Hearing_Note, 0, 0, 371, 350, "Contempt Hearing Note"
  Text 15, 10, 90, 10, "Type of Contempt Hearing"
  DropListBox 140, 5, 210, 15, "Select one:"+chr(9)+"1st Appearance"+chr(9)+"OTSC: no contempt order"+chr(9)+"Contested"+chr(9)+"OTSC: contempt order"+chr(9)+"In Custody", Hearing_Type
  Text 15, 30, 160, 10, "District Court Judge / Child Support Magistrate"
  DropListBox 175, 30, 175, 15, "Select one" +chr(9)+ child_support_magistrates_list +chr(9)+ county_judge_list, district_court_judge
  Text 15, 55, 55, 10, "County Attorney"
  DropListBox 15, 70, 145, 15, "Select one" +chr(9)+ county_attorney_list, CAO_list
  Text 180, 55, 110, 10, "Child Support Officer"
  EditBox 180, 70, 170, 15, CSO_textbox  
  CheckBox 15, 95, 55, 10, "NCP Present ", NCP_Present
  CheckBox 80, 95, 60, 10, "Pro Se", Pro_se
  Text 80, 106, 60, 10, "Represented by:"
  EditBox 145, 100, 210, 15, NCP_Represented_by
  CheckBox 15, 125, 50, 10, "CP Present", CP_Present
  Text 80, 125, 55, 10, "Represented by:"
  EditBox 145, 120, 210, 15, CP_Represented_by
  Text 15, 145, 130, 10, "Total arrears on the date of the hearing"
  EditBox 160, 140, 195, 15, Total_arrears_on_the_hearing_date
  Text 15, 165, 115, 10, "Total arrears under the contempt:"
  EditBox 135, 160, 85, 15, Total_arrears_under_the_contempt
  Text 225, 165, 20, 10, "as of"
  EditBox 250, 160, 105, 15, as_of
  Text 15, 185, 70, 10, "Summary of Hearing:"
  EditBox 100, 185, 255, 15, Summary_of_Hearing
  Text 15, 215, 65, 10, "Next Steps for CSO:"
  EditBox 100, 210, 255, 15, Next_steps_for_CSO
  Text 15, 240, 95, 10, "Next hearing date and time:"
  EditBox 110, 235, 245, 15, Next_hearing_date_and_time
  Text 15, 260, 155, 10, "Click here if a Promise to Appear was signed"
  CheckBox 160, 260, 30, 10, "", Check_Yes
  Text 15, 280, 45, 10, "Bail amount:"
  EditBox 75, 275, 85, 15, Bail_amount
  Text 170, 305, 65, 10, "Worker's Signature"
  EditBox 240, 300, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 245, 325, 50, 15
    CancelButton 305, 325, 50, 15
EndDialog

'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 85, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

call PRISM_case_number_finder(PRISM_case_number)

'Case number display dialog
Do
	Dialog case_number_dialog
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
Loop until case_number_valid = True

'Displays dialog for date of the hearing caad note and checks for information

Do
'Shows dialog, validates that PRISM is up and not timed out, with transmit
	err_msg = ""
	Dialog Contempt_Hearing_Note
	cancel_confirmation	
	CALL Prism_case_number_validation(prism_case_number, case_number_valid)
	IF worker_signature = "" THEN err_msg = err_msg & vbNEWline & "You must sign your CAAD note"
	IF Hearing_Type = "Select one:" THEN err_msg = err_msg & vbNEWline & "You must enter in a hearing type!"
	IF CAO_List = "Select one:" THEN err_msg = err_msg & vbNEWline & "You must enter CAO!"
	IF district_court_judge = "Select one:" THEN err_msg = err_msg & vbNEWline & "You must enter Judge/Magistrate!"
	IF CSO_textbox = "" THEN err_msg = err_msg & vbNEWline & "You must enter Child Support Officer!"
	IF Summary_of_Hearing = "" THEN err_msg = err_msg & vbNEWline & "You must enter hearing notes"
	IF err_msg <> "" THEN MsgBox "***Notice***" & vbNEWline & err_msg &vbNEWline & vbNEWline & "Please resolve for the script"
LOOP UNTIL err_msg = ""	


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")


'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)


PF5					'Did this because you have to add a new note

EMWriteScreen "M3909", 4, 54  'adds correct caad code 

EMSetCursor 16, 4			'Because the cursor does not default to this location

call write_bullet_and_variable_in_CAAD("Type of Contempt Hearing", Hearing_Type) 
call write_bullet_and_variable_in_CAAD("District Court Judge/Child Support Magistrate", district_court_judge)
call write_bullet_and_variable_in_CAAD("County Attorney", CAO_list)
call write_bullet_and_variable_in_CAAD("Child Support Officer", CSO_textbox)
if NCP_present = 1 then
	call write_variable_in_CAAD("* NCP present")
	call write_bullet_and_variable_in_CAAD("Represented by", NCP_Represented_by)
else 
	call write_variable_in_CAAD ("* NCP not present")
end if
if Pro_se = 1 then
	call write_variable_in_CAAD("* NCP Pro Se")
end if 
if CP_present = 1 then
	call write_variable_in_CAAD("* CP present")
	call write_bullet_and_variable_in_CAAD("Represented by", CP_Represented_by)
else 
	call write_variable_in_CAAD ("* CP not present")
end if
call write_bullet_and_variable_in_CAAD("Total arrears on the date of hearing", Total_arrears_on_the_hearing_date)
call write_bullet_and_variable_in_CAAD("Total arrears under the contempt", Total_arrears_under_the_contempt)
call write_bullet_and_variable_in_CAAD("Contempt arrears through:", as_of)
call write_bullet_and_variable_in_CAAD("Summary of Hearing", Summary_of_Hearing)
call write_bullet_and_variable_in_CAAD("Next Steps needed:", Next_steps_for_CSO)
call write_bullet_and_variable_in_CAAD("Next Hearing Date and Time", Next_hearing_date_and_time)
If Check_Yes = 1 then
	call write_variable_in_CAAD("* NCP signed promise to appear")
else
	call write_variable_in_CAAD("* NCP did not sign promise to appear")
end if

call write_bullet_and_variable_in_CAAD("Bail amount", Bail_amount)
call write_variable_in_CAAD("---")	
call write_variable_in_CAAD(worker_signature)

script_end_procedure("")
