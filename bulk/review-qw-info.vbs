'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "review-qw-info.vbs" 
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
call changelog_update("12/06/2016", "Fixed the skipping matches issue, and added spreadsheets for user reference.", "Wendy LeVesseur, Anoka County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'This is the dialog to select the CSO. The script will run off the 8 digit worker ID code entered here.
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
			
			BeginDialog select_cso_dlg, 0, 0, 286, 145, "Review Quarterly Wage Matches - select caseload"
			EditBox 70, 55, 65, 15, cso_id
			Text 70, 80, 155, 10, cso_name
			ButtonGroup ButtonPressed
				OkButton 130, 125, 50, 15
				PushButton 180, 125, 50, 15, "UPDATE CSO", update_cso_button
				PushButton 230, 125, 50, 15, "STOP SCRIPT", stop_script_button
			Text 10, 15, 265, 30, "This script will check for worklist items coded L2500 and L2501 for the following Worker ID. If you wish to change the Worker ID, enter the desired Worker ID in the box and press UPDATE CSO. When you are ready to continue, press OK."
			Text 10, 60, 50, 10, "Worker ID:"
			Text 10, 80, 55, 10, "Worker Name:"
		
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
																																				'The additional of IF ButtonPressed = -1 to the conditional statement is needed 
																																		'to allow the worker to update the CSO's worker ID without getting a warning message.
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1 
	LOOP UNTIL err_msg = ""
END FUNCTION

'=====VARIABLES TO DECLARE=====
checked = 1
unchecked = 0



'=====THE SCRIPT=====
EMConnect ""
CALL check_for_PRISM(False)

'Loading the dialog to select the CSO
CALL select_cso(ButtonPressed, cso_id, cso_name)

msgbox "This script will mark as reviewed quarterly wage matches for NCP that already exist on NCID.  Then it will mark as reviewed quarterly wage matches for CP that already exist on CPID. " &_
	"The script produces spreadsheets of data for reference. Please click OK to continue."

'First we are going to process L2500 (NCP's worklists)
CALL write_value_and_transmit("L2500", 20, 30)

uswt_row = 7
DO
	EMReadScreen uswt_type_id, 5, uswt_row, 45
	EMReadScreen prism_case_number, 13, uswt_row, 8
	prism_case_number = replace(prism_case_number, " ", "-")
	IF uswt_type_id = "L2500" THEN cases_array_ncp = cases_array_ncp & prism_case_number & " "
	uswt_row = uswt_row + 1
	IF uswt_row = 19 THEN 
		PF8
		uswt_row = 7
	END IF
LOOP UNTIL uswt_type_id <> "L2500" 

cases_array_ncp = trim(cases_array_ncp)
cases_array_ncp = split(cases_array_ncp, " ")

number_of_cases = ubound(cases_array_ncp)
DIM ncp_array()
ReDim ncp_array(number_of_cases, 6)

'>>>> HERE ARE THE 6 POSITIONS WITHIN THE ARRAY <<<<
'ncp_array(i, 0) >> PRISM_case_number
'ncp_array(i, 1) >> NCP Name
'ncp_array(i, 2) >> Employer on Summary Screen
'ncp_array(i, 3) >> Employer on Wage Match
'ncp_array(i, 4) >> Known employer? (1 for Yes, 0 for No)
'ncp_array(i, 5) >> Purged? (1 for Yes, 0 for No)

position_number = 0
FOR EACH prism_case_number IN cases_array_ncp
'	ncp_array(i, 0) >> PRISM_case_number
	IF prism_case_number <> "" THEN 
		ncp_array(position_number, 0) = prism_case_number
		position_number = position_number + 1
	END IF
NEXT


placeholder_case_number_string = ""
FOR i = 0 to number_of_cases
'ncp_array(i, 0) >> PRISM_case_number
'ncp_array(i, 1) >> NCP Name
'ncp_array(i, 2) >> Employer on Summary Screen
'ncp_array(i, 3) >> Employer on Wage Match
'ncp_array(i, 4) >> Known employer? (1 for Yes, 0 for No)
'ncp_array(i, 5) >> Purged? (1 for Yes, 0 for No)
	CALL navigate_to_PRISM_screen("NCSU")
	EMWriteScreen ncp_array(i, 0), 4, 8
	EMWriteScreen right(ncp_array(i, 0), 2), 4, 19
	CALL write_value_and_transmit("D", 3, 29)

	'Storing NCP's name into the array.
 	EMReadScreen ncp_name, 30, 7, 12
	ncp_name = trim(ncp_name)
	ncp_array(i, 1) = ncp_name	
	
	'Storing NCP's employer from the summary screen into the array.
	EMReadScreen NCID_emp, 30, 13, 49
	ncp_array(i, 2) = NCID_emp

'If the script does not find this case number in the placeholder string, we will build on that string,
'and we will go ahead with the logic to check the quarterly wage match on this case.
	
	CALL navigate_to_PRISM_screen("NCQW")
	
		QW_row = 9
		placeholder_qw_string = ""

			DO 
			'Exit do if "End of Data" is read.					
				EMReadScreen end_of_data_check, 11, QW_row, 32
				IF end_of_data_check = "End of Data" THEN EXIT DO	
		
				'When an unreviewed result is found, need to display it.
				EMReadScreen rev_check, 1, QW_row, 75
				IF rev_check <> "Y" THEN
					EMReadScreen qw_string, 65, QW_row, 8
					IF InStr(placeholder_qw_string, qw_string) = 0 THEN
						placeholder_qw_string = placeholder_qw_string & "~~~" & qw_string
				
						EMWriteScreen "D", QW_row, 4
						transmit

						'Obtain wage match employer information and store in the array.
						EMReadScreen wage_match_employer, 30, 9, 12
						ncp_array(i, 3) = wage_match_employer
						PF6

						'If employer already exists on employer screen, mark the case reviewed.						
						EMReadScreen bottom_line_message, 70, 24, 3
						bottom_line_message = trim(bottom_line_message) 'Check for error messages on the bottom of the screen
						IF bottom_line_message <> "" THEN 
							IF InStr(bottom_line_message, "already exists") <> 0 THEN 'This employer is known
							PF3
							EMWriteScreen "M", 3, 29  	'Modify the page
							EMWriteScreen "Y", 16, 64     'Mark reviewed
							ncp_array(i, 4) = 1           'Indicator that this employer is known
							ncp_array(i, 5) = 1           'Indicator that worklist should be purged (if PRISM doesn't purge it automatically when it is marked reviewed)
							transmit
							PF3   'return to the qw screen
							ELSEIF InStr(bottom_line_message, "pf6 to select") <> 0 THEN 'Error message indicates the user must select location for the employer.
							PF3									 'This worklist will be left unreviewed for the user to process manually.
							PF3
							ELSEIF InStr(bottom_line_message, "Fein is required") <> 0 THEN 'Error message indicates the FEIN is unavailable.
							PF3									    'This worklist will be left unreviewed for the user to process manually.
							ELSE 'Some other message is displayed.  This worklist will be left unreviewed for the user to process manually.
							PF3
							PF3
							END IF
						ELSEIF bottom_line_message = "" THEN 'This worklist will be left unreviewed for the user to process manually.
							PF3
							PF3
											
						END IF
					END IF			
				END IF			
				QW_row = QW_row + 1
				IF QW_row = 19 THEN      	'Pagination
					PF8
					QW_row = 9
				END IF
			
			LOOP UNTIL end_of_data_check = "End of Data"
	
NEXT

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

objExcel.Cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "NCP NAME"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "EMPLOYER FROM SUMMARY SCREEN"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "EMPLOYER FROM WAGE MATCH"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "KNOWN EMPLOYER?"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "PURGED?"
objExcel.Cells(1, 6).Font.Bold = True

excel_row = 2

'Updating the Excel spreadsheet with information
FOR i = 0 to number_of_cases 
	FOR k = 0 to 6
		objExcel.Cells(excel_row, k + 1).Value = ncp_array(i, k)
		IF k = 5 or k = 4 THEN 
			IF ncp_array(i, k) = checked THEN 
				objExcel.Cells(excel_row, k + 1).Value = "Y"		
			END IF
			IF ncp_array(i, k) = unchecked THEN 
				objExcel.Cells(excel_row, k + 1).Value = "N"
			END IF
		END IF
	NEXT
	excel_row = excel_row + 1
NEXT

'Autofitting each column.
FOR x_col = 1 to 6
	objExcel.Columns(x_col).AutoFit()
NEXT

'One more check for PRISM
CALL check_for_PRISM(False)

'Purging the worklists that are marked to purge.  Usually this is done by PRISM when the worklist is marked reviewed, but not always....
number_of_cases_purged = 0
FOR i = 0 to number_of_cases

	IF ncp_array(i, 5) = checked THEN 
		CALL navigate_to_PRISM_screen("CAWT")
		CALL write_value_and_transmit("L2500", 20, 29)
		EMWriteScreen left(ncp_array(i, 0), 10), 20, 8	
		EMWritescreen right(ncp_array(i, 0), 2), 20, 19
		transmit
		
	
		DO
			EMReadscreen cawd_type, 5, 8, 8
			IF cawd_type = "L2500" THEN
				EMWriteScreen "P", 8, 4
				transmit
				transmit
			END IF
		LOOP UNTIL cawd_type <> "L2500"
	END IF
NEXT

'---------------------------------------------------------------------------------------------------------------------------------------------
'Now, basically doing the same steps over again, this time for L2501 (CP's worklists)

CALL navigate_to_PRISM_screen("USWT")
CALL write_value_and_transmit("L2501", 20, 30)

uswt_row = 7
DO
	EMReadScreen uswt_type_id, 5, uswt_row, 45
	EMReadScreen prism_case_number, 13, uswt_row, 8
	prism_case_number = replace(prism_case_number, " ", "-")
	IF uswt_type_id = "L2501" THEN cases_array_cp = cases_array_cp & prism_case_number & " "
	uswt_row = uswt_row + 1
	IF uswt_row = 19 THEN 
		PF8
		uswt_row = 7
	END IF
LOOP UNTIL uswt_type_id <> "L2501" 

cases_array_cp = trim(cases_array_cp)
cases_array_cp = split(cases_array_cp, " ")

number_of_cases = ubound(cases_array_cp)
DIM cp_array()
ReDim cp_array(number_of_cases, 6)

'>>>> HERE ARE THE 6 POSITIONS WITHIN THE ARRAY <<<<
'cp_array(i, 0) >> PRISM_case_number
'cp_array(i, 1) >> CP Name
'cp_array(i, 2) >> Employer on Summary Screen
'cp_array(i, 3) >> Employer on Wage Match
'cp_array(i, 4) >> Known employer? (1 for Yes, 0 for No)
'cp_array(i, 5) >> Purged? (1 for Yes, 0 for No)

position_number = 0
FOR EACH prism_case_number IN cases_array_cp
'	cp_array(i, 0) >> PRISM_case_number
	IF prism_case_number <> "" THEN 
		cp_array(position_number, 0) = prism_case_number
		position_number = position_number + 1
	END IF
NEXT


placeholder_case_number_string = ""
FOR i = 0 to number_of_cases
'cp_array(i, 0) >> PRISM_case_number
'cp_array(i, 1) >> CP Name
'cp_array(i, 2) >> Employer on Summary Screen
'cp_array(i, 3) >> Employer on Wage Match
'cp_array(i, 4) >> Known employer?(1 for Yes, 0 for No)
'cp_array(i, 5) >> Purged? (1 for Yes, 0 for No)
	CALL navigate_to_PRISM_screen("CPSU")
	EMWriteScreen cp_array(i, 0), 4, 8
	EMWriteScreen right(cp_array(i, 0), 2), 4, 19
	CALL write_value_and_transmit("D", 3, 29)

	'Storing CP's name into the array. 
	EMReadScreen cp_name, 30, 6, 12
	cp_name = trim(cp_name)
	cp_array(i, 1) = cp_name	
	
	'Storing CP's employer from the summary screen into the array.
	EMReadScreen CPID_emp, 30, 11, 12
	cp_array(i, 2) = CPID_emp

'If the script does not find this case number in the placeholder string, we will build on that string,
'and we will go ahead with the logic to check the quarterly wage match on this case.
	
	CALL navigate_to_PRISM_screen("CPQW")
	
		QW_row = 9
		placeholder_qw_string = ""
		qw_string = ""
			DO 		
			'Exit do if "End of Data" is read.		
				EMReadScreen end_of_data_check, 11, QW_row, 32 
				IF end_of_data_check = "End of Data" THEN EXIT DO	
		
				'When an unreviewed result is found, need to display it.
				EMReadScreen rev_check, 1, QW_row, 75
				IF rev_check <> "Y" THEN
					EMReadScreen qw_string, 65, QW_row, 8
					IF InStr(placeholder_qw_string, qw_string) = 0 THEN
						placeholder_qw_string = placeholder_qw_string & "~~~" & qw_string
					
						EMWriteScreen "D", QW_row, 4
						transmit

						'Obtain wage match employer information and store in the array.
						EMReadScreen wage_match_employer, 30, 9, 12
						cp_array(i, 3) = wage_match_employer
						PF6

						'If employer already exists on employer screen, mark the case reviewed.					
						EMReadScreen bottom_line_message, 70, 24, 3
						bottom_line_message = trim(bottom_line_message) 'Check for error messages on the bottom of the screen
						IF bottom_line_message <> "" THEN 
							IF InStr(bottom_line_message, "already exists") <> 0 THEN 'This employer is known
								PF3
								EMWriteScreen "M", 3, 29  	'Modify the page
								EMWriteScreen "Y", 16, 64     'Mark reviewed
								cp_array(i, 4) = 1			'Indicator that the employer is known
								cp_array(i, 5) = 1			'Indicator that the worklist should be purged (if PRISM doesn't purge it automatically when it is marked reviewed)
								transmit
								PF3   'return to the qw screen
							ELSEIF InStr(bottom_line_message, "pf6 to select") <> 0 THEN 'Error message indicates the user must select location for the employer.
								PF3														 'This worklist will be left unreviewed for the user to process manually.
								PF3
							ELSEIF InStr(bottom_line_message, "Fein is required") <> 0 THEN 'Error message indicates the FEIN number is unavailable.
								PF3															'This worklist will be left unreviewed for the user to process manually.				
							ELSE 'Some other message is displayed.  This worklist will be left unreviewed for the user to process manually.
								PF3
								PF3
							END IF
						ELSEIF bottom_line_message = "" THEN 'This worklist will be left unreviewed for the user to process manually
							PF3
							PF3	
						END IF
					END IF
				END IF			
				QW_row = QW_row + 1
				IF QW_row = 19 THEN      	'Pagination
					PF8
					QW_row = 9
				END IF
			
			LOOP UNTIL end_of_data_check = "End of Data"
NEXT


Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

objExcel.Cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "CP NAME"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "EMPLOYER FROM SUMMARY SCREEN"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "EMPLOYER FROM WAGE MATCH"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "KNOWN EMPLOYER?"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "PURGED?"
objExcel.Cells(1, 6).Font.Bold = True

excel_row = 2

'Updating the Excel spreadsheet information.
FOR i = 0 to number_of_cases 
	FOR k = 0 to 6
		objExcel.Cells(excel_row, k + 1).Value = cp_array(i, k)
		IF k = 5 or k = 4 THEN 
			IF cp_array(i, k) = checked THEN 
				objExcel.Cells(excel_row, k + 1).Value = "Y"		
			END IF
			IF cp_array(i, k) = unchecked THEN 
				objExcel.Cells(excel_row, k + 1).Value = "N"
			END IF
		END IF
	NEXT
	excel_row = excel_row + 1
NEXT

'Autofitting each column.
FOR x_col = 1 to 6
	objExcel.Columns(x_col).AutoFit()
NEXT

'One more check for PRISM
CALL check_for_PRISM(False)

'Purging the worklists that are marked to purge.  Usually this is done by PRISM when the worklist is marked reviewed, but not always....

FOR i = 0 to number_of_cases

	IF cp_array(i, 5) = checked THEN 
		CALL navigate_to_PRISM_screen("CAWT")
		CALL write_value_and_transmit("L2501", 20, 29)
		EMWriteScreen left(cp_array(i, 0), 10), 20, 8	
		EMWritescreen right(cp_array(i, 0), 2), 20, 19
		transmit
		
	
		DO
			EMReadscreen cawd_type, 5, 8, 8
			IF cawd_type = "L2501" THEN
				EMWriteScreen "P", 8, 4
				transmit
				transmit
			END IF
		LOOP UNTIL cawd_type <> "L2501"
	END IF
NEXT

script_end_procedure("Success!! The script is now ending")

