'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "county-attorney-referral.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denomination = "C"
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


'Calling dialog details pop up box'

BeginDialog CAO_REFERRAL, 0, 0, 366, 170, "County Attorney Referral"
  Text 10, 10, 60, 10, "CAAD Note Type"
  DropListBox 70, 10, 280, 15, "M2653 DOCUMENTS TO ATTORNEY FOR APPROVAL"+chr(9)+"M2655 DOCUMENTS APPROVED BY ATTORNEY"+chr(9)+"O5168 COUNTY ATTORNEY CONTACT REGARDING LEGAL ACTION"+chr(9)+"O4700 DOCS SENT TO CTY ATTORNEY FOR REVIEW AND SIGNATURE"+chr(9)+"O4701 DOCS RETURNED FROM COUNTY ATTORNEY WITH SIGNATURE"+chr(9)+"O4702 DOCS RETURNED FROM CTY ATTORNEY WITHOUT SIGNATURE ", CAAD_Note_Type_dropdown
  Text 10, 40, 50, 10, "Case Number"
  EditBox 70, 35, 140, 15, PRISM_case_number
  Text 10, 60, 130, 10, "Attorney Referral was submitten to:"
  DropListBox 140, 60, 165, 15, "Select one:" +chr(9)+ county_attorney_list, CAO_list
  Text 10, 95, 65, 10, "Question/Response:"
  EditBox 85, 90, 270, 15, Question_box
  Text 10, 125, 80, 10, "Worker's Signature"
  EditBox 85, 120, 270, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 245, 145, 50, 15
    CancelButton 305, 145, 50, 15
EndDialog


'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

'Searches for the case number.
CALL PRISM_case_number_finder (PRISM_case_number)

'Displays dialog for Modification caad note and checks for information

Do
'Shows dialog, validates PRISM mandated fields completed, with transmit
	err_msg = ""
	Dialog CAO_REFERRAL
	cancel_confirmation
	CALL Prism_case_number_validation(prism_case_number, case_number_valid)
	IF CAAD_Note_Type_dropdown = "" THEN err_msg = err_msg & vbNEWline & "You must select a CAAD note type!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNEWline & "You must sign your CAAD note!"
	IF CAO_List = "" THEN err_msg = err_msg & vbNEWline & "You must enter who the referral was submitted to!"
	IF Question_box = "" THEN err_msg = err_msg & vbNEWline & "You must enter why referral submitted!"
	IF err_msg <> "" THEN MsgBox "***Notice***" & vbNEWline & err_msg &vbNEWline & vbNEWline & "Please resolve for the script"
LOOP UNTIL err_msg = ""


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")


'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)


PF5					'Did this because you have to add a new note

EMWriteScreen Left(CAAD_Note_Type_dropdown, 5), 4, 54  'adds correct caad code

EMSetCursor 16, 4			'Because the cursor does not default to this location

call write_editbox_in_PRISM_case_note("Attorney sent to", CAO_list, 4)
call write_editbox_in_PRISM_case_note("Question/Response", Question_box, 4)
call write_new_line_in_PRISM_case_note(worker_signature)

script_end_procedure("")
