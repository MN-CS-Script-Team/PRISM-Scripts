'Collecting stats---------------------------------------------------------------
script_name = "BULK - EVALUATE NOCS.vbs"
start_time = timer

'These variables need to be dimmed to work properly with the custom functions in the script.
DIM initial_run_through
DIM worker_signature
DIM cso_name

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


'=====DIALOGS=====
BeginDialog confirm_dlg, 0, 0, 281, 135, "Confirmation"
  ButtonGroup ButtonPressed
    PushButton 220, 65, 50, 15, "RETRY", retry_button
    OkButton 220, 110, 50, 15
  Text 10, 40, 260, 20, "If you feel any of the cases need revision, press RETRY. The script will then run through all the cases one more time."
  Text 10, 10, 260, 20, "You have reviewed all the cases. You can double check script's work using the existing Excel file."
  Text 10, 90, 265, 10, "Otherwise, if you think the cases are ready for processing, press OK to continue."
EndDialog

'=====ADDITIONAL CUSTOM FUNCTIONS=====

'This first custom function is used to build the case review dialog. The values pulled through are "i" and "nocs_array."
'Changes in "i" change the case number and case information being worked on.
'nocs_array holds the information for that case.
FUNCTION build_NOCS_dlg(i, nocs_array)

	'nocs_array(i, 7) >> The specific DORD doc to be sent. Because the dialog uses check boxes, the incoming value has to be converted from text
	'to a 1/0 value in the check boxes.
	IF nocs_array(i, 7) = "F0111" THEN
		f0111_checkbox = 1
		f0115_checkbox = 0
	ELSEIF nocs_array(i, 7) = "F0115" THEN
		f0111_checkbox = 0
		f0115_checkbox = 1
	ELSEIF nocs_array(i, 7) = "NONE" THEN
		f0111_checkbox = 0
		f0115_checkbox = 0
	END IF


	BeginDialog NOCS_dlg, 0, 0, 311, 330, "Notice of Continued Service"
	  Text 70, 15, 105, 10, nocs_array(i, 0) 'Prism case number
	  Text 10, 70, 50, 10, "Program Code:"
	  Text 65, 70, 55, 10, nocs_array(i, 3) 'Program code
	  Text 215, 70, 50, 10, "Full Services?"
	  Text 270, 70, 55, 10, nocs_array(i, 4) 'Full services indicator
	  Text 120, 70, 65, 10, "IV-D Cooperation?"
	  Text 190, 70, 10, 10, nocs_array(i, 6) 'IV-D cooperation code
	  Text 10, 95, 110, 10, "Program Code Effective Date:"
	  EditBox 120, 90, 80, 15, nocs_array(i, 5) ' Effective date of program change from CATH
	  CheckBox 95, 130, 165, 10, "F0111 - NPA Notice of Continued Services", f0111_checkbox
	  CheckBox 95, 145, 165, 10, "F0115 - MA/MCRE Notice of Continued Services", f0115_checkbox
	  CheckBox 10, 170, 155, 10, "Check HERE to generate a FREE worklist", nocs_array(i, 8)
	  EditBox 75, 185, 55, 15, nocs_array(i, 9) 'FREE worklist date
	  EditBox 115, 205, 190, 15, nocs_array(i, 10) 'FREE worklist notes
	  EditBox 120, 235, 185, 15, nocs_array(i, 12) 'CAAD note for Other Notes
	  CheckBox 10, 265, 265, 10, "Check HERE to PURGE the D0800 Evaluate for Continued Services worklist.", nocs_array(i, 11) 'Purge worklist indicator
	  EditBox 135, 285, 80, 15, nocs_array(i, 13) 		'<< NOT worker_signature ... that is being used elsewhere and this needs to be separate.
	  	ButtonGroup ButtonPressed
	  	OkButton 150, 310, 50, 15
		PushButton 200, 310, 50, 15, "SKIP CASE", skip_case_button
	  	PushButton 250, 310, 55, 15, "STOP SCRIPT", stopscript_button
	  Text 10, 15, 50, 10, "Case Number:"
	  Text 10, 130, 70, 10, "DORD Doc to Send:"
	  Text 25, 190, 50, 10, "Worklist Date:"
	  Text 25, 210, 85, 10, "Additional Worklist Text:"
	  Text 10, 290, 125, 10, "I authorize that this work is accurate."
	  Text 20, 30, 65, 10, "Custodial Parent:"
	  Text 85, 30, 90, 10, nocs_array(i, 1) 'CP Name
	  Text 20, 45, 75, 10, "Non-Custodial Parent:"
	  Text 100, 45, 90, 10, nocs_array(i, 2) 'NCP Name
	  Text 10, 240, 105, 10, "Additional Notes for CAAD Note: "
	EndDialog

	'We load the dialog within the function for each case.
	'The do/loop allows for error checking.
	DO
		'reseting the value for err_msg
		err_msg = ""
		'loading the dialog
		DIALOG NOCS_dlg
			IF ButtonPressed = stopscript_button THEN script_end_procedure("The script has stopped.")		'If the user presses the "STOP SCRIPT" button, the script stops
			IF ButtonPressed = -1 AND f0111_checkbox = 1 AND f0115_checkbox = 1 THEN err_msg = err_msg & vbCr & "* You cannot send both the F0111 and the F0115. Please pick one or none."		'If the user selects the F0111 and F0115 docs to be send, err_msg is updated. Both docs cannot be generated.
			IF (nocs_array(i, 8) = 1 AND IsDate(nocs_array(i, 9)) = False) THEN err_msg = err_msg & vbCr & "* You must format the Worklist Date as a date."					'If the user checks to generate a FREE worklist item but does not enter the date as a date, err_msg is updated.
			IF nocs_array(i, 13) = "" AND ButtonPressed <> skip_case_button THEN err_msg = err_msg & vbCr & "* You must authorize the work as accurate. Please sign off on your work."			'If the user tries to move off the case without using the "SKIP CASE" button (using the "OK" button), err_msg is updated.
			IF ButtonPressed = -1 AND err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbcr & err_msg & vbCr & vbCr & "You must resolve for the script to continue."			'IF err_msg is not blank (meaning one or more of the error conditions above is/are met) the script will display the error(s) to the user.

			IF ButtonPressed = skip_case_button AND initial_run_through = True THEN 		'IF the user elects to skip this case and this is the first time through...
				nocs_array(i, 7) = "NONE"		'Do not send a DORD doc
				nocs_array(i, 8) = 0			'Do not create FREE worklist
				nocs_array(i, 11) = 0			'Do not purge
				nocs_array(i, 12) = ""			'Do not CAAD Note
			END IF											'...the value initial_run_through is important. We do not want the user to work on several cases, run decide to review the cases again, and have to duplicate their work.
															'... If the user is double checking their work (meaning this is NOT the first time they are running through the dialogs), pressing "SKIP CASE" will not reset the values in the dialog.
	LOOP UNTIL ButtonPressed = skip_case_button OR (ButtonPressed = -1 AND err_msg = "")		'In order to get out of the do/loop, the user must press either the "SKIP CASE" button, OR press "OK" and have a null err_msg


	'Converting the value for the DORD doc to be sent BACK to human speak.
	IF nocs_array(i, 7) <> "NONE" THEN
		IF f0111_checkbox = 1 THEN nocs_array(i, 7) = "F0111"
		IF f0115_checkbox = 1 THEN nocs_array(i, 7) = "F0115"
		IF f0111_checkbox = 0 AND f0115_checkbox = 0 THEN nocs_array(i, 7) = "NONE"
	END IF

END FUNCTION

'	This is the dialog to select the CSO. The script will run off the 8 digit worker ID code entered here.
FUNCTION select_cso(ButtonPressed, cso_id, cso_name)
	DO
		DO
			CALL navigate_to_PRISM_screen("USWT")
			err_msg = ""
			'Grabbing the CSO name for the intro dialog.
			CALL find_variable("Worker Id: ", cso_id, 8)
			EMSetCursor 20, 13
			PF1
			CALL write_value_and_transmit(cso_id, 20, 35)
			EMReadScreen cso_name, 24, 13, 55
			cso_name = trim(cso_name)
			PF3

			BeginDialog select_cso_dlg, 0, 0, 286, 145, "Notice of Continued Service - Select CSO"
			EditBox 70, 55, 65, 15, cso_id
			Text 70, 80, 155, 10, cso_name
			EditBox 115, 100, 75, 15, worker_signature
			ButtonGroup ButtonPressed
				OkButton 130, 125, 50, 15
				PushButton 180, 125, 50, 15, "UPDATE CSO", update_cso_button
				PushButton 230, 125, 50, 15, "STOP SCRIPT", stop_script_button
			Text 10, 15, 265, 30, "This script will check for worklist items coded D0800 for the following Worker ID. If you wish to change the Worker ID, enter the desired Worker ID in the box and press UPDATE CSO. When you are ready to continue, press OK."
			Text 10, 60, 50, 10, "Worker ID:"
			Text 10, 80, 55, 10, "Worker Name:"
			Text 10, 105, 100, 10, "Please sign your CAAD notes:"
			EndDialog

			DIALOG select_cso_dlg
				IF ButtonPressed = stop_script_button THEN script_end_procedure("The script has stopped.")
				IF ButtonPressed = update_cso_button THEN
					CALL navigate_to_PRISM_screen("USWT")
					CALL write_value_and_transmit(cso_id, 20, 13)
					EMReadScreen cso_name, 24, 13, 55
					cso_name = trim(cso_name)
				END IF
				IF cso_id = "" THEN err_msg = err_msg & vbCr & "* You must enter a Worker ID."
				IF len(cso_id) <> 8 THEN err_msg = err_msg & vbCr & "* You must enter a valid, 8-digit Worker ID."
				IF (ButtonPressed = -1 AND worker_signature = "") THEN err_msg = err_msg & vbCr & "* Please sign your CAAD note."		'<< If the worker tries to continue without signing, the warning bell will sound.
																																		'The additional of IF ButtonPressed = -1 to the conditional statement is needed
																																		'to allow the worker to update the CSO's worker ID without getting a warning message.
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1
	LOOP UNTIL err_msg = ""
END FUNCTION

'=====THE SCRIPT=====
EMConnect ""
CALL check_for_PRISM(False)

'Loading the dialog to select the CSO
CALL select_cso(ButtonPressed, cso_id, cso_name)

'And away we go...
CALL write_value_and_transmit("D0800", 20, 30)

uswt_row = 7
DO
	EMReadScreen uswt_type_id, 5, uswt_row, 45
	EMReadScreen prism_case_number, 13, uswt_row, 8
	prism_case_number = replace(prism_case_number, " ", "-")
	IF uswt_type_id = "D0800" THEN cases_array = cases_array & prism_case_number & " "
	uswt_row = uswt_row + 1
	IF uswt_row = 19 THEN
		PF8
		uswt_row = 7
	END IF
LOOP UNTIL uswt_type_id <> "D0800"

cases_array = trim(cases_array)
cases_array = split(cases_array, " ")

number_of_cases = ubound(cases_array)
DIM nocs_array()
ReDim nocs_array(number_of_cases, 13)

'>>>> HERE ARE THE 14 positions within the array <<<<
'nocs_array(i, 0) >> PRISM_case_number
'nocs_array(i, 1) >> CP name
'nocs_array(i, 2) >> NCP name
'nocs_array(i, 3) >> program_code (NPA, non-NPA; pulled from CALI)
'nocs_array(i, 4) >> full_service (whether or not the case is full service; pulled from CAST)
'nocs_array(i, 5) >> program type change date; pulled from CATH
'nocs_array(i, 6) >> IV-D Cooperation Code
'nocs_array(i, 7) >> DORD doc to send
'nocs_array(i, 8) >> generate a "FREE" worklist ("Y" or "N")
'nocs_array(i, 9) >> FREE worklist date
'nocs_array(i, 10) >> FREE worklist notes
'nocs_array(i, 11) >> Purge? (1 for Yes, 0 for No)
'nocs_array(i, 12) >> CAAD note "Other Notes"
'nocs_array(i, 13) >> worker authorization..."I authorize this information is correct."" This is part of the multidimensional array for the do-looping. We do not want the worker to have to re-authorize work.

position_number = 0
FOR EACH prism_case_number IN cases_array
'	nocs_array(i, 0) >> PRISM_case_number
	IF prism_case_number <> "" THEN
		nocs_array(position_number, 0) = prism_case_number
		position_number = position_number + 1
	END IF
NEXT


CALL navigate_to_PRISM_screen("CAST")
FOR i = 0 to number_of_cases
'	nocs_array(i, 0) >> PRISM_case_number
'	nocs_array(i, 1) >> CP name
'	nocs_array(i, 2) >> NCP name
'	nocs_array(i, 3) >> program_code (NPA, non-NPA)
'	nocs_array(i, 4) >> full_service (whether or not the case is full service; pulled from CAST)
	EMWriteScreen nocs_array(i, 0), 4, 8
	EMWriteScreen right(nocs_array(i, 0), 2), 4, 19
	CALL write_value_and_transmit("D", 3, 29)
	EMReadScreen full_service, 1, 9, 60
	nocs_array(i, 4) = full_service
	EMReadScreen cp_name, 35, 6, 12
	EMReadScreen ncp_name, 35, 7, 12
	cp_name = trim(cp_name)
	ncp_name = trim(ncp_name)
	nocs_array(i, 1) = cp_name
	nocs_array(i, 2) = ncp_name
	EMReadScreen program_code, 3, 6, 68
	nocs_array(i, 3) = program_code
NEXT

CALL navigate_to_PRISM_screen("CATH")
FOR i = 0 to number_of_cases
'	nocs_array(i, 0) >> PRISM_case_number
'	nocs_array(i, 5) >> program type change date; pulled from CATH
	EMWriteScreen nocs_array(i, 0), 20, 8
	EMWriteScreen right(nocs_array(i, 0), 2), 20, 19
	transmit
	PF7
	cath_row = 8
	DO
		EMReadScreen case_program_code, 17, cath_row, 11
		IF case_program_code <> "CASE PROGRAM CODE" THEN
			cath_row = cath_row + 1
			IF cath_row = 20 THEN
				PF8
				cath_row = 8
			END IF
		ELSEIF case_program_code = "CASE PROGRAM CODE" THEN
			EMReadScreen program_change_date, 8, cath_row, 2
			date_converter_PALC_PAPL (program_change_date)
			nocs_array(i, 5) = program_change_date

			EXIT DO
		END IF
	LOOP
NEXT

CALL navigate_to_PRISM_screen("GCSC")
FOR i = 0 to number_of_cases
'	nocs_array(i, 0) >> PRISM_case_number
'	nocs_array(i, 3) >> program_code (NPA, non-NPA; pulled from CALI)
'	nocs_array(i, 4) >> full_service (whether or not the case is full service; pulled from CAST)
'	nocs_array(i, 6) >> IV-D Cooperation Code
'	nocs_array(i, 7) >> DORD doc to send
'	nocs_array(i, 11) >> Purge? (1 for Yes, 0 for No)
	EMWriteScreen nocs_array(i, 0), 4, 8
	EMWriteScreen right(nocs_array(i, 0), 2), 4, 19
	current_date = date_converter_PALC_PAPL(date)
	EMWriteScreen current_date, 9, 18
	CALL write_value_and_transmit("D", 3, 29)
	EMReadScreen ivd_coop_code, 1, 15, 25
	IF ivd_coop_code = "_" THEN ivd_coop_code = "Y"			'IF there has never been non-coop for good cause, the panel will be coded "_" which is effectively a "Y"
	nocs_array(i, 6) = ivd_coop_code
	'Default PURGE value = False
	nocs_array(i, 11) = 0

	IF nocs_array(i, 3) = "NPA" THEN			'IF the case is NPA THEN
		IF nocs_array(i, 4) = "Y" THEN 			'IF the case is FULL SERVICE
			nocs_array(i, 7) = "F0111"				'Send the F0111
		ELSEIF nocs_array(i, 4) <> "Y" THEN 	'IF the case is NOT FULL SERVICE
			nocs_array(i, 7) = "F0115"				'Send the F0115
			IF nocs_array(i, 6) = "N" THEN nocs_array(i, 8) = 1	'IF the case is in Good Cause NON-COOP, you need to create a free worklist item
		END IF
	END IF
NEXT

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

objExcel.Cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "CUSTODIAL PARENT"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "NON-CUSTODIAL PARENT"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "PROGRAM CODE"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "FULL SERVICE?"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "PROGRAM CHANGE DATE"
objExcel.Cells(1, 6).Font.Bold = True
objExcel.Cells(1, 7).Value = "GOOD CAUSE COOP?"
objExcel.Cells(1, 7).Font.Bold = True
objExcel.Cells(1, 8).Value = "DORD TO SEND"
objExcel.Cells(1, 8).Font.Bold = True
objExcel.Cells(1, 9).Value = "FREE WORKLIST?"
objExcel.Cells(1, 9).Font.Bold = True
objExcel.Cells(1, 10).Value = "FREE WORKLIST DATE"
objExcel.Cells(1, 10).Font.Bold = True
objExcel.Cells(1, 11).Value = "WORKLIST NOTES"
objExcel.Cells(1, 11).Font.Bold = True
objExcel.Cells(1, 12).Value = "PURGE?"
objExcel.Cells(1, 12).Font.Bold = True

'Updating the Excel spreadsheet with initial information
FOR i = 0 to number_of_cases
	FOR j = 0 to 11
		objExcel.Cells(i + 2, j + 1).Value = nocs_array(i, j)
	NEXT
NEXT

'Autofitting each column.
FOR x_col = 1 to 11
	objExcel.Columns(x_col).AutoFit()
NEXT

initial_run_through = True
DO
	'Running the dialog for each case.
	excel_row = 2
	FOR i = 0 to number_of_cases
		'This changes the back fill of the selected row to draw the worker's attention.
		FOR select_column = 1 to 12
			objExcel.Cells(excel_row, select_column).Interior.ColorIndex = 6		'Setting the background fill to yellow
			objExcel.Cells(excel_row - 1, select_column).Interior.ColorIndex = 2	'Setting the previous background fill to white
		NEXT
		'Building the dialog
		CALL build_NOCS_dlg(i, nocs_array)

		'Updating the Excel spreadsheet in real time.
		FOR j = 0 to 11
			objExcel.Cells(excel_row, j + 1).Value = nocs_array(i, j)
			IF j = 8 THEN
				IF nocs_array(i, 8) = 1 THEN objExcel.Cells(excel_row, j + 1).Value = "Y"
				IF nocs_array(i, 8) = 0 THEN objExcel.Cells(excel_row, j + 1).Value = "N"
			END IF
			IF j = 11 THEN
				IF nocs_array(i, j) = 1 THEN objExcel.Cells(excel_row, j + 1).Value = "Y"
				IF nocs_array(i, j) = 0 THEN objExcel.Cells(excel_row, j + 1).Value = "N"
			END IF
		NEXT
		excel_row = excel_row + 1
	NEXT

	FOR select_column = 1 to 12
		objExcel.Cells(excel_row - 1, select_column).Interior.ColorIndex = 2
	NEXT

	DIALOG confirm_dlg
		IF ButtonPressed <> -1 THEN initial_run_through = False
LOOP UNTIL ButtonPressed = -1

'Redoing the autofit for the columns.
FOR x_col = 1 to 11
	objExcel.Columns(x_col).AutoFit()
NEXT

'One more check for PRISM
CALL check_for_PRISM(False)

'Converting numeric values assigned by checkboxes to True or False.
FOR i = 0 to number_of_cases
'	nocs_array(i, 11) >> Purge? (1 for Yes, 0 for No)
	IF nocs_array(i, 11) = 1 THEN
		nocs_array(i, 11) = True
	ELSE
		nocs_array(i, 11) = False
	END IF
NEXT

'The script will now generate a CAAD note...
CALL navigate_to_PRISM_screen("CAAD")
FOR i = 0 to number_of_cases
'	nocs_array(i, 12) >> CAAD note "Other Notes"
	IF nocs_array(i, 12) <> "" THEN
		EMWriteScreen left(nocs_array(i, 0), 10), 20, 8
		EMWriteScreen right(nocs_array(i, 0), 2), 20, 19
		transmit

		PF5
		EMWriteScreen "FREE", 4, 54
		EMSetCursor 16, 4
		CALL write_variable_in_CAAD("*** EVALUATED FOR NOTICE OF CONTINUED SERVICE ***")
		CALL write_bullet_and_variable_in_CAAD("Worker Notes", nocs_array(i, 12))
		CALL write_variable_in_CAAD("---")
		CALL write_variable_in_CAAD(worker_signature)
		CALL write_variable_in_CAAD("~~Generated using automated script.")
		transmit
		PF3
	END IF
NEXT

'The script will now generate the appropriate DORD docs..
CALL navigate_to_PRISM_screen("DORD")
FOR i = 0 to number_of_cases
	EMWriteScreen nocs_array(i, 0), 4, 15
	EMWriteScreen right(nocs_array(i, 0), 2), 4, 26
	transmit
	IF nocs_array(i, 7) <> "NONE" THEN
		CALL write_value_and_transmit("A", 3, 29)
		EMWriteScreen "        ", 4, 50
		EMWriteScreen "       ", 4, 59
		CALL write_value_and_transmit(nocs_array(i, 7), 6, 36)
		PF14
		CALL write_value_and_transmit("U", 20, 14)
		CALL write_value_and_transmit("S", 7, 5)
		CALL create_mainframe_friendly_date(nocs_array(i, 5), 16, 15, "YYYY")
		transmit
		PF3
		CALL write_value_and_transmit("M", 3, 29)
	END IF
NEXT

'Now the script needs generate the FREE worklist and insert the free text
CALL navigate_to_PRISM_screen("USWT")
FOR i = 0 to number_of_cases
'	nocs_array(i, 2) >> NCP name
'	nocs_array(i, 3) >> program_code (NPA, non-NPA; pulled from CALI)
'	nocs_array(i, 4) >> full_service (whether or not the case is full service; pulled from CAST)
'	nocs_array(i, 6) >> IV-D Cooperation Code
'	nocs_array(i, 8) >> generate a "FREE" worklist ("Y" or "N")
'	nocs_array(i, 9) >> FREE worklist date
'	nocs_array(i, 10) >> FREE worklist notes
	IF nocs_array(i, 8) = 1 THEN
		PF5
		EMWriteScreen left(nocs_array(i, 0), 10), 4, 8
		EMWriteScreen right(nocs_array(i, 0), 2), 4, 19
		EMWriteScreen "FREE", 4, 37
		CALL create_mainframe_friendly_date(nocs_array(i, 9), 17, 21, "YYYY")
		EMSetCursor 10, 4
		IF nocs_array(i, 3) <> "NPA" AND nocs_array(i, 4) = "N" AND nocs_array(i, 6) = "N" THEN CALL write_variable_in_CAAD("* Review this case for possible closure. Case flipped to NPA, client is not full service ,and GCSC shows case is IV-D Non-Coop.")
		CALL write_variable_in_CAAD(nocs_array(i, 10))
		CALL write_variable_in_CAAD("~~Generated using automated script.")
		transmit
		PF3
	END IF
NEXT

'Now the script needs to PURGE for all (i, 11) = True

number_of_cases_purged = 0
FOR i = 0 to number_of_cases
'	nocs_array(i, 0) >> PRISM_case_number
'	nocs_array(i, 11) >> Purge? (1 for Yes, 0 for No)
	IF nocs_array(i, 11) = True THEN
		CALL navigate_to_PRISM_screen("CAWT")
		CALL write_value_and_transmit("D0800", 20, 29)
		EMWriteScreen left(nocs_array(i, 0), 10), 20, 8
		EMWritescreen right(nocs_array(i, 0), 2), 20, 19
		transmit


		DO
			EMReadscreen cawd_type, 5, 8, 8
			IF cawd_type = "D0800" THEN
				EMWriteScreen "P", 8, 4
				transmit
				transmit
				number_of_cases_purged = number_of_cases_purged + 1
			END IF
		LOOP UNTIL cawd_type <> "D0800"
	END IF
NEXT

script_end_procedure("Success!! " &  number_of_cases_purged  & " items have been purged.")
