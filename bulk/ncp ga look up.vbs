
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "list-generator---ncp ga look up.vbs"
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
call changelog_update("10/19/2017", "Initial version", "Wendy LeVesseur for Robert Kalb, Anoka County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

' adding a message box to notify the user that they are identified as a supervisor
IF supervisor_user = true THEN 
	run_sup_mode = MsgBox ("Supervisory User Detected" & vbCr & _
					"ID: " & UCASE(windows_user_id) & vbCr & _
					vbCr & _
					"You are able to run this script in ''supervisor mode.'' This run mode will enables two additional features..." & vbCr & _
					vbTab & "1. You can run this script for an entire unit (TEAM), rather than a single position. To do this, enter a value for ''TEAM'' and leave ''POSITION'' blank." & vbCr & _
					vbTab & "2. The script will ask if you want to add worklist items to CAWT for the cases identified with an NCP active on GA or MFIP." & vbCr & _
					vbCr & _
					"Do you wish to proceed in ''supervisor mode?''", vbYesNo, vbInformation)
	IF run_sup_mode = vbYes THEN
		supervisor_mode = TRUE
	ELSEIF run_sup_mode = vbNo THEN 
		supervisor_mode = FALSE
	END IF
END IF

' >>>>> The Dialog <<<<<
BeginDialog CALI_selection_dialog, 0, 0, 236, 100, "NCP GA Look Up"
  EditBox 140, 30, 25, 15, cali_team
  EditBox 205, 30, 25, 15, cali_position
  Text 5, 10, 205, 10, "Enter these fields to run this script on another CALI caseload:"
  Text 5, 35, 55, 10, "County:  " & county_cali_code
  Text 60, 35, 45, 10, "Office:  001"
  Text 110, 35, 25, 10, "Team:"
  Text 170, 35, 30, 10, "Position:"
  ButtonGroup ButtonPressed
    OkButton 125, 80, 50, 15
    CancelButton 180, 80, 50, 15
  IF supervisor_mode = true then Text 55, 60, 230, 20, "*** SUPERVISOR MODE ENABLED ***"
EndDialog

'***********************************************************************************************************************************************
'If the user is already on the CALI screen when the script is run, results may be inaccurate.  Also, if the user runs the script when the
'position listing screen is open, the screen must be exited before the script can run properly.  This function checks to see if either of
'these circumstances apply.  If the position list is open, the script exits the list, and if the CALI screen is open, navigates away so that
'the report will function properly.
FUNCTION refresh_CALI_screen
	EMReadScreen check_for_position_list, 22, 8, 36
		IF check_for_position_list = "Caseload Position List" THEN
			PF3
		END IF
	EMReadScreen check_for_caseload_list, 13, 2, 32
		If check_for_caseload_list = "Caseload List" THEN
			CALL navigate_to_PRISM_screen("MAIN")
			transmit
		END IF
END FUNCTION

'Connects to Bluezone
EMConnect ""

' checking that we are not timed out of PRISM...
check_for_PRISM(TRUE)

'loading the dialog
DO
	' err_msg handling
	err_msg = ""
	Dialog CALI_selection_dialog
		IF ButtonPressed = 0 THEN StopScript
		IF cali_team = "" THEN err_msg = err_msg & vbCr & "* CALI Team field is blank."
		IF supervisor_mode = false AND cali_position = "" THEN err_msg = err_msg & vbCr & "* CALI Position field is blank."
		IF len(cali_team) <> 3 THEN err_msg = err_msg & vbCr & "* The length of CALI Team must be 3 characters."
		IF cali_position <> "" AND len(cali_position) <> 2 THEN err_msg = err_msg & vbCr & "* The length of CALI Position must be 2 characters."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

' checking again that we are not timed out of PRISM...
check_for_PRISM(false)

'this script has 3 possible run modes...
'	1. Supervisor Mode, Entire Team
'	2. Supervisor Mode, Single CALI
'	3. CSO Mode

'	when the user is a supervisor, they can run for an entire team. this will be enabled when the CS
IF supervisor_mode = true AND cali_position = "" THEN 
	cali_array = ""
	
	' Resetting PRISM...
	CALL navigate_to_PRISM_screen("REGL")
	transmit
	
	' Going to CALI...
	CALL navigate_to_PRISM_screen("CALI")
	EMSetCursor 20, 40
	PF1
	
	' writing the county code and office 001  
	EMWriteScreen county_cali_code, 20, 22
	EMWriteScreen "001", 20, 34
	
	' writing the team number and transmitting to start searching for positions that match that value...
	CALL write_value_and_transmit(cali_team, 20, 44)
	
	' setting a variable for the row we are searching
	cali_search_row = 13
	DO
		' checking for 'End of data'
		EMReadScreen end_of_data_check, 70, cali_search_row, 3
		IF InStr(end_of_data_check, "End of Data") <> 0 THEN EXIT DO
		
		' creating, reading, and assigning a value to a variable that will be used as a comparison with the original CALI team
		EMReadScreen cali_search_team, 3, cali_search_row, 31
		IF cali_search_team = cali_team THEN 			' if they match, read the cali position and throw it all into the array
			EMReadScreen cali_search_position, 20, cali_search_row, 18
			cali_search_position = replace(cali_search_position, " ", "")
			cali_array = cali_array & cali_search_position & ","
		END IF
		
		' next row
		cali_search_row = cali_search_row + 1
		IF cali_search_row = 19 THEN 
			cali_search_row = 13
			PF8
		END IF
	LOOP
	
	' when we are done, add END to it, remove '',END'' and split by comma
	cali_array = cali_array & "END"
	cali_array = replace(cali_array, ",END", "")
	cali_array = split(cali_array, ",")
	
	PF3
ELSE
	' else, treat the single CALI position as an array for to navigate the next bit
	cali_array = county_cali_code & "001" & cali_team & cali_position
	cali_array = split(cali_array)
END IF

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True 
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

' column headers
ObjExcel.Cells(1, 1).Value = "PRISM CASE NUMBER"
objExcel.Cells(1, 2).Value = "NCP MCI"
ObjExcel.Cells(1, 3).Value = "NCP NAME"
objExcel.Cells(1, 4).Value = "NCP SSN"
objExcel.Cells(1, 5).Value = "NCP PUBLIC ASSISTANCE STATUS"

'sets row to fill info into Excel
excel_row = 2

FOR EACH cali_user IN cali_array
	CALL navigate_to_PRISM_screen("CALI")  'Navigate to CALI, remove any case number entered, and display the desired CALI listing
	EMWriteScreen "             ", 20, 58
	EMWriteScreen "  ", 20, 69
	
	EMSetCursor 20, 18
	EMSendKey cali_user
	transmit

	prism_row = 8

	'setting the script on the NCP area of CALI...
	DO
		PF11
		EMReadScreen search_for_ncp_name_on_CALI, 8, 6, 35
	LOOP UNTIL search_for_ncp_name_on_CALI = "NCP Name"

	Do 
		'Loops script until the end of CALI
		'Copies Case Number, Function Type, Program Type, CP Name, and NCP Name to the Excel document
		
		EmReadscreen end_of_data_check, 70, prism_row, 3
		If InStr(end_of_data_check, "End of Data") <> 0 then exit do
		
		EMReadScreen prism_case_number, 14, prism_row, 7 'Reads and copies case number
			prism_case_number = replace(prism_case_number, "  ", "-")
		EMReadScreen NCP_MCI, 10, prism_row, 22 'reads the NCP's MCI
		EMReadScreen NCP_name, 26, prism_row, 33 'Reads and copies NCP name
	
		NCP_MCI = Cstr(NCP_MCI)
		
		'Set rows in Excel for case number, funtion type, program type, CP name, and NCP name
		ObjExcel.Cells(excel_row, 1).Value = prism_case_number
		objExcel.Cells(excel_row, 2).Value = NCP_MCI & " MCI"		'adding an arbitrary character so to as preserve the leading zeros
		ObjExcel.Cells(excel_row, 3).Value = NCP_name
	
		prism_row = prism_row + 1
		excel_row = excel_row + 1
	
		IF prism_row = 19 THEN
			PF8
			prism_row = 8
		END IF
	Loop Until end_of_data_check = "End of Data"
NEXT

' grabbing the NCP's social security number off NCDE
excel_row = 2
CALL navigate_to_PRISM_screen("NCDE")
DO
	ncp_mci = objExcel.Cells(excel_row, 2).Value
	ncp_name = objExcel.Cells(excel_row, 3).Value
	
	'accounting for instances where the CSO does not have access to the case
	IF InStr(ncp_name, "Access Denied") = 0 THEN 
		EMWriteScreen ncp_mci, 4, 7
		CALL write_value_and_transmit("D", 3, 29)
		
		EMReadScreen ncp_ssn, 11, 10, 7
		ncp_ssn = replace(ncp_ssn, " ", "-")
	
		objExcel.Cells(excel_row, 4).Value = ncp_ssn
	ELSE
		objExcel.Cells(excel_row, 4).Value = "___-__-____"
	END IF
	
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""

MsgBox "The script has finished gathering data from PRISM and is ready to start checking for MAXIS. Press OK for the script to continue."

DO
	'creating a variable for to find MAXIS
	maxis_found = false
	
	' using ASCII table as a sub for A, B... chr(65) = "A"
	FOR mx_sesh = 65 TO 69
		EMConnect ( chr(mx_sesh) )
		EMReadScreen found_maxis, 5, 1, 39
		IF found_maxis = "MAXIS" OR found_maxis = "AXIS " THEN 
			EMFocus 	' bringing that BZ session to the front
			maxis_found = true
			EXIT FOR
		END IF
	NEXT
	
	IF maxis_found = false THEN 
		no_maxis_found = MsgBox("The script could not find an active MAXIS session." & vbCr & vbCr & _ 
								"Please activate MAXIS in BlueZone session S1, S2, S3, S4, or S5 and then press ''OK'' the script to continue." & vbCr & vbCr & _
								"If you would like the stop the script, press ''CANCEL.''", vbOKCancel + vbInformation)
		IF no_maxis_found = vbCancel THEN script_end_procedure("Script cancelled.")
	END IF
LOOP UNTIL maxis_found = true

'checking that MAXIS is not timed out...sorry about that, Pam...
CALL check_for_MAXIS(false)

excel_row = 2
DO
	back_to_SELF
	CALL write_value_and_transmit("PERS", 16, 43)
	client_ssn = objExcel.Cells(excel_row, 4).Value
	
	IF client_ssn <> "___-__-____" THEN 
		ssn_prefix = left(client_ssn, 3)
		ssn_mid = right(left(client_ssn, 6), 2)
		ssn_end = right(client_ssn, 4)
		
		EMWriteScreen ssn_prefix, 14, 36
		EMWriteScreen ssn_mid, 14, 40
		CALL write_value_and_transmit(ssn_end, 14, 43)
		
		'checking to see that this client exists in MAXIS
		EMReadScreen at_pers, 4, 2, 47
		IF at_pers <> "PERS" THEN 
			'if we got off PERS, then we are on DSPL
			'entering GA in the Program Selection
			CALL write_value_and_transmit("GA", 7, 22)
			
			EMReadScreen ga_current, 7, 10, 35
			EMReadScreen ga_appl, 1, 10, 61
			IF ga_current = "Current" AND ga_appl = "Y" THEN 
				EMReadScreen ga_start_date, 8, 10, 25
				EMReadScreen maxis_case_number, 8, 10, 6
				objExcel.Cells(excel_row, 5).Value = "NCP active on General Assistance on MAXIS case " & trim(maxis_case_number) & " since " & ga_start_date & "."
			ELSE
				CALL write_value_and_transmit("MF", 7, 22)
				
				EMReadScreen mf_current, 7, 10, 35
				EMReadScreen mf_appl, 1, 10, 61
				IF mf_current = "Current" THEN 
					EMReadScreen maxis_case_number, 8, 10, 6
					maxis_case_number = trim(maxis_case_number)
					IF mf_appl = "Y" and mf_current = "Current" THEN 
						EMReadScreen mf_start_date, 8, 10, 25
						objExcel.Cells(excel_row, 5).Value = "NCP active on MFIP on MAXIS case " & trim(maxis_case_number) & " since " & mf_start_date & "."
					ELSE						
						'navigating to MEMB to grab the client's ref num
						CALL navigate_to_MAXIS_screen("STAT", "MEMB")
						EMReadScreen at_self, 4, 2, 50
						IF at_self = "SELF" THEN 
							'privileged MAXIS case
						ELSE
							memb_ref_num = ""
							DO
								EMReadScreen memb_ssn, 11, 7, 42
								memb_ssn = replace(memb_ssn, " ", "-")
								IF memb_ssn <> objExcel.Cells(excel_row, 4).Value THEN 
									transmit
									EMReadScreen enter_a_valid, 13, 24, 2
									IF enter_a_valid = "ENTER A VALID" THEN 
										client_found = false
										EXIT DO
									END IF									
								ELSE
									EMReadScreen memb_ref_num, 2, 4, 33
									client_found = true
									EXIT DO
								END IF
							LOOP
							
							IF client_found = true THEN 
								CALL navigate_to_MAXIS_screen("ELIG", "MFIP")
							
								elig_row = 7
								DO
									EMReadScreen elig_ref_num, 2, elig_row, 6
									IF elig_ref_num = memb_ref_num THEN 
										EMReadScreen elig_status, 4, elig_row, 53
										IF elig_status = "ELIG" THEN objExcel.Cells(excel_row, 5).Value = "NCP active on MFIP on MAXIS case " & trim(maxis_case_number) & " since " & mf_start_date & "."
										exit do
									ELSEIF elig_ref_num <> memb_ref_num THEN 
										elig_row = elig_row + 1
									ELSEIF elig_ref_num = "  " THEN 
										exit do
									END IF
								LOOP
							END IF
						END IF
					END IF
				END IF
			END IF	
		END IF
	END IF
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""

IF supervisor_mode = TRUE THEN 
	'going back to PRISM to create worklist items

	create_worklist = MsgBox ("The script has finished gathering data from MAXIS. Do you want it to return to PRISM to create worklists?", vbYesNo, vbInformation)

	IF create_worklist = vbYes THEN 
	
		DO
			'creating a variable for to find PRISM
			PRISM_found = false
			
			' using ASCII table as a sub for A, B... chr(65) = "A"
			FOR prism_sesh = 65 TO 69
				EMConnect ( chr(prism_sesh) )
				EMReadScreen found_PRISM, 5, 1, 36
				IF found_PRISM = "PRISM" THEN 
					EMFocus 	' bringing that BZ session to the front
					PRISM_found = true
					EXIT FOR
				END IF
			NEXT
		
			IF PRISM_found = false THEN 
				no_PRISM_found = MsgBox("The script could not find an active PRISM session." & vbCr & vbCr & _ 
										"Please activate PRISM in BlueZone session S1, S2, S3, S4, or S5 and then press ''OK'' the script to continue." & vbCr & vbCr & _
										"If you would like the stop the script, press ''CANCEL.''", vbOKCancel + vbInformation)
				IF no_PRISM_found = vbCancel THEN script_end_procedure("Script cancelled.")
			END IF
		LOOP UNTIL PRISM_found = true
		
		CALL check_for_PRISM(false)
	
		excel_row = 2
	
		CALL navigate_to_PRISM_screen("CAWT")
		PF5
	
		DO
			prism_case_number = objExcel.Cells(excel_row, 1).Value
			public_assistance_status = objExcel.Cells(excel_row, 5).Value
		
			IF public_assistance_status <> "" THEN 
				EMWriteScreen "A", 3, 30
				EMWriteScreen left(prism_case_number, 10), 4, 8
				EMWriteScreen right(prism_case_number, 2), 4, 19

				EMWriteScreen "FREE", 4, 37
				EMSetCursor 10, 4
				EMSendKey public_assistance_status

				transmit
			END IF
			
			excel_row = excel_row + 1
		LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""
	END IF
END IF

'Autofitting columns
For col_to_autofit = 1 to 5
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

script_end_procedure("Success!! The script has finished running.")

