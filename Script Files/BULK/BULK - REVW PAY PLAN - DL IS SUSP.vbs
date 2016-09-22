'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - E4111 SUSP SCRUBBER.vbs"
start_time = timer

Dim URL, REQ, FSO					'Declares variables to be good to option explicit users

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


' >>>>> THE SCRIPT <<<<<
EMConnect ""
CALL check_for_PRISM(False)

'Loading the dialog to select the CSO
CALL select_cso_caseload(ButtonPressed, cso_id, cso_name)

'And away we go...
Call navigate_to_PRISM_screen ("USWT")
CALL write_value_and_transmit("E4111", 20, 30)

uswt_row = 7
DO
	EMReadScreen uswt_type_id, 5, uswt_row, 45
	EMReadScreen prism_case_number, 13, uswt_row, 8
	prism_case_number = replace(prism_case_number, " ", "-")
	IF uswt_type_id = "E4111" THEN cases_array = cases_array & prism_case_number & " "
	uswt_row = uswt_row + 1
	IF uswt_row = 19 THEN
		PF8
		uswt_row = 7
	END IF
LOOP UNTIL uswt_type_id <> "E4111"

cases_array = trim(cases_array)
cases_array = split(cases_array, " ")

number_of_cases = ubound(cases_array)
DIM info_array()
ReDim info_array(number_of_cases, 5)

'>>>> HERE ARE THE positions within the array <<<<
'info_array(i, 0) >> PRISM_case_number
'info_array(i, 1) >> CP name
'info_array(i, 2) >> NCP name
'info_array(i, 3) >> NCP MCI
'info_array(i, 4) >> DL Suspended?
'info_array(i, 5) >> Purge?

position_number = 0
FOR EACH prism_case_number IN cases_array
'	info_array(i, 0) >> PRISM_case_number
	IF prism_case_number <> "" THEN
		info_array(position_number, 0) = prism_case_number
		position_number = position_number + 1
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
objExcel.Cells(1, 4).Value = "NCP MCI?"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "DL SUSPENDED?"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "PURGE?"
objExcel.Cells(1, 6).Font.Bold = True
'Updating the Excel spreadsheet with initial information
FOR i = 0 to number_of_cases
	FOR j = 0 to 5
		objExcel.Cells(i + 2, j + 1).Value = info_array(i, j)
	NEXT
NEXT

'Autofitting each column.
FOR x_col = 1 to 5
	objExcel.Columns(x_col).AutoFit()
NEXT

CALL navigate_to_PRISM_screen("CAST")
FOR i = 0 to number_of_cases
'info_array(i, 0) >> PRISM_case_number
'info_array(i, 1) >> CP name
'info_array(i, 2) >> NCP name
'info_array(i, 3) >> NCP MCI

	EMWriteScreen info_array(i, 0), 4, 8
	EMWriteScreen right(info_array(i, 0), 2), 4, 19
	CALL write_value_and_transmit("D", 3, 29)
	EMReadScreen cp_name, 35, 6, 12
	EMReadScreen ncp_name, 35, 7, 12
	cp_name = trim(cp_name)
	ncp_name = trim(ncp_name)
	info_array(i, 1) = cp_name
	info_array(i, 2) = ncp_name
	EMReadScreen ncp_mci, 10, 9, 11
	info_array(i, 3) = ncp_mci
NEXT

CALL navigate_to_PRISM_screen("ENFL")
FOR i = 0 to number_of_cases
'info_array(i, 4) >> DL Suspended?
	EMWriteScreen info_array(i, 3), 20, 7
	EMWriteScreen "DLS", 20, 43
	transmit
	ENFL_row = 8
	DO
		EMReadScreen end_of_data, 11, ENFL_row, 32
		IF end_of_data <> "End of Data" THEN
			EMReadScreen ENFL_status, 3, ENFL_row, 9
			EMReadScreen ENFL_case_no, 12, ENFL_row, 67
			case_number = replace(info_array(i, 0), "-", "")
			IF USWT_row = 19 THEN
				PF8
				USWT_row = 8
			END IF
			IF ENFL_status = "ACT" AND ENFL_case_no = case_number THEN
				info_array(i, 4) = TRUE
				info_array(i, 5) = TRUE

				EXIT DO
			END IF
			IF ENFL_status <> "ACT" AND ENFL_case_no = case_number THEN
				info_array(i, 4) = FALSE
				info_array(i, 5) = FALSE

				EXIT DO
			END IF
			ENFL_row = ENFL_row + 1
		END IF
	LOOP UNTIL end_of_data = "End of Data"

NEXT
'Updating the Excel spreadsheet with initial information
FOR i = 0 to number_of_cases
	FOR j = 0 to 5
		objExcel.Cells(i + 2, j + 1).Value = info_array(i, j)
	NEXT
NEXT

number_of_worklists_purged = 0
FOR i = 0 to number_of_cases
'	info_array(i, 0) >> PRISM_case_number
'	info_array(i, 5) >> Purge?
	IF info_array(i, 5) = True THEN
		CALL navigate_to_PRISM_screen("CAWT")
		CALL write_value_and_transmit("E4111", 20, 29)
		EMWriteScreen left(info_array(i, 0), 10), 20, 8
		EMWritescreen right(info_array(i, 0), 2), 20, 19
		transmit

		DO
			EMReadscreen cawd_type, 5, 8, 8
			IF cawd_type = "E4111" THEN
				EMWriteScreen "P", 8, 4
				transmit
				transmit
				number_of_worklists_purged = number_of_worklists_purged + 1
			END IF
		LOOP UNTIL cawd_type <> "E4111"
	END IF
NEXT


script_end_procedure("Success!  " & number_of_worklists_purged & " worklists purged!")
