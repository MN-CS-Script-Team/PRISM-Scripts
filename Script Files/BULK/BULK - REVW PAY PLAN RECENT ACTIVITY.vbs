'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - E4111 WORKLIST SCRUBBER.vbs"
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

' >>>>> THE SCRIPT <<<<<
EMConnect ""

'>>>>> GOING TO USWT <<<<<
Call navigate_to_Prism_screen("USWT")

' >>>>> SELECTING THE SPECIFIC WORKLIST TYPE <<<<<
EMWriteScreen "E4111", 20, 30
transmit

ENFL_row = 8
USWT_row = 7
count = 0
SCROLL = 0

' >>>>> STARTING THE DO LOOP. THE SCRIPT NEEDS TO HANDLE THESE CASES ONE AT A TIME <<<<<
DO
		'reset variables
	total_voluntary_alloc = 0
	total_involuntary_alloc = 0

	EMReadScreen USWT_type, 5, USWT_row, 45
	IF USWT_type = "E4111" THEN
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		EMWriteScreen "s", USWT_row, 4
		transmit
		'Selecting the worklist brings the user to NCP's PAPD screen
		EMWriteScreen "B", 3, 29
		transmit
		EMSetCursor 8, 39
		transmit
		EMReadScreen curr_pmt, 13, 13, 16
		EMReadScreen arrs_pmt, 13, 14, 16
		curr_pmt = replace(curr_pmt, "_", "0")
		arrs_pmt = replace(arrs_pmt, "_", "0")

		pay_plan_pmt = ccur(curr_pmt) + ccur(arrs_pmt)

		Call navigate_to_PRISM_screen ("PALC")
		current_month_minus1 = DateAdd("m", -1, date)
	'	current_month = DateAdd("m", 0, date)
	'	current_day_minus1 = DateAdd("d", -1, current_day_minus1)
		c_month = datepart("m", current_month_minus1)
			IF len(c_month) = 1 THEN c_month = "0" & c_month
		c_begin_date = c_month & "/01/" & datepart("yyyy", current_month_minus1)
		begin_date = cdate(c_begin_date)
		begin_date_plus1 = DateAdd("m", 1, begin_date)
		end_date = DateAdd("d", -1, begin_date_plus1)
		CALL create_mainframe_friendly_date(c_begin_date, 20, 35, "YYYY")
		CALL create_mainframe_friendly_date(end_date, 20, 49, "YYYY")
		transmit


row = 9		'Setting variable for the do...loop

Do
	EMReadScreen end_of_data_check, 19, row, 28									'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do							'Exits do if we have

	'Reading payment date, which for some crazy reason is YYMMDD, without slashes. This converts.
	EMReadScreen pmt_ID_YY, 2, row, 7
	EMReadScreen pmt_ID_MM, 2, row, 9
	EMReadScreen pmt_ID_DD, 2, row, 11
	pmt_ID_date = pmt_ID_MM & "/" & pmt_ID_DD & "/" & pmt_ID_YY

		EMReadScreen proc_type, 3, row, 25														'Reading the proc type
		EMReadScreen case_alloc_amt, 10, row, 70
		EMReadScreen payment_type, 1, row, 55 										'Reading the amt allocated
		If payment_type = "I" then               ' check to make sure the payment status is identified, not refunded
			If proc_type = "FTS" or proc_type = "MCE" or proc_type = "NOC" or proc_type = "IFC" or proc_type = "OST" or _
			proc_type = "PCA" or proc_type = "PIF" or proc_type = "STJ" or proc_type = "STS" or proc_type = "FTJ" then 		'If proc type is one of these, it's involuntary. Else, it's voluntary.
				total_involuntary_alloc = total_involuntary_alloc + ccur(case_alloc_amt)							'Adds the alloc amt for involuntary
			Else
				total_voluntary_alloc = total_voluntary_alloc + ccur(case_alloc_amt)							'Adds the alloc amt for voluntary)
			End if
		End if

	row = row + 1														'Increases the row variable by one, to check the next row
	EMReadScreen end_of_data_check, 19, row, 28									'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do							'Exits do if we have
	If row = 19 then														'Resets row and PF8s
		PF8
		row = 9
	End if
Loop until end_of_data_check = "*** End of Data ***"

If total_involuntary_alloc = "" then total_involuntary_alloc = "0"
If total_voluntary_alloc = "" then total_voluntary_alloc = "0"

string_for_msgbox = "---PAYMENT BREAKDOWN FOR " & begin_date & " THROUGH " & end_date & "---" & chr(10) & chr(10) & "Involuntary: " & FormatCurrency(total_involuntary_alloc) & chr(10) & "Voluntary: "_
		 & FormatCurrency(total_voluntary_alloc) & chr(10) & chr(10) & "Current Pay Plan Monthly Amount (current + arrears): " & FormatCurrency(pay_plan_pmt)& chr(10) & chr(10) &_
		"PURGE THIS WORKLIST?"

purge_box = Msgbox(string_for_msgbox, 3,  "Purge this worklist?")

purge = false 'Reset the purge variable
If purge_box = "2" then stopscript  'user clicked cancel
If purge_box = "6" then purge = true  'user clicked yes
If purge_box = "7" then purge = false	'user clicked no


		' >>>>> CHECKING THE INFORMATION ON ENFL <<<<
		EMReadScreen end_of_data, 11, ENFL_row, 32
		IF end_of_data <> "End of Data" THEN

			' >>>>> READING THE STATUS AND CASE NUMBER <<<<<
			EMReadScreen ENFL_status, 3, ENFL_row, 9
			EMReadScreen ENFL_case_no, 12, ENFL_row, 67
			trimmed_case_number = Replace(USWT_case_number, " ", "", 1)
				If ENFL_status = "ACT" then
					If ENFL_case_no = trimmed_case_number then
						purge = True
						Msgbox "E4111 worklist for " & ENFL_case_no & " will be purged."
					End If
				End If

		END IF


		Call navigate_to_PRISM_screen ("CAWT")
		EMWriteScreen "E4111", 20, 29
		EMWriteScreen USWT_case_number, 20, 8
		transmit

		' >>>>> IF THE WORKLIST ITEM IS ELIGIBLE TO BE PURGED, THE SCRIPT PURGES...
		IF purge = True THEN
			cawt_row = 8
			Do
				EMReadScreen cawd_type, 5, cawt_row, 8
				if cawd_type = "E4111" then
					EMWriteScreen "P", cawt_row, 4
					transmit
					transmit
					PF3
				end if
				count = count + 1
				cawt_row = cawt_row +1
			Loop until cawd_type <> "E4111"
			Call navigate_to_PRISM_screen ("USWT")

			EMWriteScreen "E4111", 20, 30
			transmit
			IF SCROLL > 0 THEN
				FOR I = 0 TO SCROLL
				PF8
				NEXT
			END IF
			USWT_row = USWT_row + 1
			IF USWT_row = 19 THEN
				PF8
				USWT_row = 7
				SCROLL = SCROLL + 1
			END IF


		ELSE
		'  ...  IF THE WORKLIST ITEM IS NOT ELIGIBLE TO BE PURGED, THE SCRIPT INCREASES USWT_ROW + 1 <<<<<
		' >>>>> GOING BACK TO USWT <<<<<
			Call navigate_to_PRISM_screen ("USWT")

			EMWriteScreen "E4111", 20, 30
			transmit
			IF SCROLL > 0 THEN
				FOR I = 0 TO SCROLL
				PF8
				NEXT
			END IF
			USWT_row = USWT_row + 1
			IF USWT_row = 19 THEN
				PF8
				USWT_row = 7
				SCROLL = SCROLL + 1
			END IF

		END IF
	END IF

LOOP UNTIL USWT_type <> "E4111"


script_end_procedure("Success!  " & count & " worklists purged!")
