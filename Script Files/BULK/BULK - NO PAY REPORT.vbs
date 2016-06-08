'Gathering Stats--------------------------------------------------------------------
name_of_script = "BULK - NO PAY REPORT.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'Building the dialog
BeginDialog arrears_dialog, 0, 0, 276, 130, "Collections Report"
  EditBox 55, 15, 205, 15, position_array
  DropListBox 65, 75, 65, 10, "Select one..."+chr(9)+"30 days"+chr(9)+"60 days"+chr(9)+"90 days"+chr(9)+"120 days"+chr(9)+"6 months"+chr(9)+"1 year"+chr(9)+"2 years"+chr(9)+"No payment", minimum_date_range
  DropListBox 65, 100, 65, 15, "Open Ended"+chr(9)+"60 days"+chr(9)+"90 days"+chr(9)+"120 days"+chr(9)+"6 months"+chr(9)+"1 year"+chr(9)+"2 years", maximum_date_range
  CheckBox 150, 70, 120, 10, "Check here for ARREARS ONLY", arrears_only_checkbox
  ButtonGroup ButtonPressed
    OkButton 170, 105, 45, 15
    CancelButton 215, 105, 45, 15
  Text 10, 20, 40, 10, "Position(s):"
  GroupBox 10, 60, 130, 65, "Date Range"
  Text 20, 80, 40, 10, "More than: "
  Text 20, 105, 40, 10, "Less than: "
  Text 10, 35, 260, 15, "**NOTE: You can run this report for multiple users if you separate their position numbers with a comma."
EndDialog


'***********************************************************************************************************************************************
'This is a custom function to populate data from CALI into the Excel spreadsheet.  The function lists the caseload worker ID, PRISM Case #,
'and CSO name for the Child Support Unit indicated by the index parameter.

FUNCTION write_CALI_data_in_excel (index)

	EMWriteScreen ("CS" & index), 20, 40
	transmit
	FOR j = 0 to 99
		IF len(j) <> 2 THEN j = "0" & j
			EMWriteScreen j, 20, 49
			transmit
			EMReadScreen error_msg, 20, 24, 2
			error_msg = trim (error_msg)
			IF error_msg = "" THEN
				EMReadScreen arr_only, 8, 4, 13
				
				IF inStr(arr_only, "ARR") <> 0 THEN   'if the caseload worker id includes "ARR" then do the following:
					cali_row = 8  'navigates to the first case listed in CALI 
						DO
							EMReadScreen end_of_data, 11, cali_row, 32    ' Goes through all the cases from CALI and writes them to an excel spreadsheet
							IF end_of_data <> "End of Data" THEN
								EMReadScreen PRISM_case_number, 14, cali_row, 7
								PRISM_case_number = replace(PRISM_case_number, "  ", "-")
								EMReadScreen cso_name, 30, 4, 28
								cso_name = trim(cso_name)
								objExcel.Cells(excel_row, 1).Value = arr_only 'Worker ID
								objExcel.Cells(excel_row, 2).Value = PRISM_case_number 'PRISM Case #
								objExcel.Cells(excel_row, 5).Value = cso_name 'CSO name
								excel_row = excel_row + 1
								cali_row = cali_row + 1
							END IF
							IF cali_row = 19 THEN    'Navigate to a new page
								cali_row = 8
								PF8
							END IF
						LOOP UNTIL end_of_data = "End of Data"
				END IF
			END IF
	NEXT
END FUNCTION
'***********************************************************************************************************************************************
'If the user is already on the CALI screen when the script is run, report may be inaccurate.  Also, if the user runs the script when the 
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
'***********************************************************************************************************************************************

'Determining the start of the fiscal year
IF datepart("M", date) < 10 THEN
	fiscal_start = "10/01/" & datepart("YYYY", dateadd("YYYY", -1, date))
ELSE
	fiscal_start = "10/01/" & datepart("YYYY", date)
END IF

EMConnect ""

CALL check_for_PRISM(true)
IF county_cali_code = "" THEN script_end_procedure("Your agency is not properly configured for this script. Please refer this error message to your scripts administrator.")

'Grabbing user ID to validate user of script. Only some users are allowed to use all of the functionality of this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

DO
	DO
		worker_array = ""
		err_msg = ""
		dialog arrears_dialog
			IF ButtonPressed = 0 THEN stopscript
			
			'Adding data validation to the list of users we are going to check with this script...
			position_array = replace(position_array, " ", "")
			IF position_array = "" THEN 
				err_msg = err_msg & vbCr & "* You must enter at least 1 PRISM position."
			ELSE
				IF InStr(position_array, ",") <> 0 THEN
					position_array = split(position_array, ",")
					FOR EACH cali_position IN position_array
						worker_not_found = false
						IF cali_position <> "" THEN 
							IF len(cali_position) = 8 THEN
								CALL navigate_to_PRISM_screen("CALI")
								CALL navigate_to_PRISM_screen("REGL")
								transmit
								CALL write_value_and_transmit(county_cali_code, 20, 18)
								CALL EMSetCursor(20, 49)
								PF1
								cali_search_row = 13
								DO
									CALL EMReadScreen(worker_position, 8, cali_search_row, 39)
									IF UCASE(worker_position) <> UCASE(cali_position) THEN 
										cali_search_row = cali_search_row + 1
										IF cali_search_row = 19 THEN 
											cali_search_row = 13
											PF8
										END IF
										CALL EMReadScreen(end_of_data, 11, cali_search_row, 39)
										IF UCASE(end_of_data) = "END OF DATA" THEN 
											worker_not_found = TRUE
											err_msg = err_msg & vbCr & "* Worker at position " & cali_position & " was not found."
											EXIT DO
										END IF
									ELSE
										CALL EMReadScreen(full_position, 20, cali_search_row, 18)
										full_position = replace(full_position, " ", "")
										worker_array = worker_array & full_position & ","
										EXIT DO
									END IF
								LOOP
							ELSEIF len(cali_position) = 11 THEN 
								worker_array = worker_array & cali_position & ","
							ELSEIF len(cali_position) <> 8 AND len(cali_position) <> 11 THEN 
								err_msg = err_msg & vbCr & "* CALI position " & cali_position & " is not valid."
							END IF
						END IF
						PF3
					NEXT
				ELSEIF InStr(position_array, ",") = 0 THEN 
					IF len(position_array) = 8 THEN
						CALL navigate_to_PRISM_screen("CALI")
						CALL navigate_to_PRISM_screen("REGL")
						transmit
						CALL write_value_and_transmit(county_cali_code, 20, 18)
						CALL EMSetCursor(20, 49)
						PF1
						cali_search_row = 13
						DO
							CALL EMReadScreen(worker_position, 8, cali_search_row, 39)
							IF UCASE(worker_position) <> UCASE(position_array) THEN 
								cali_search_row = cali_search_row + 1
								IF cali_search_row = 19 THEN 
									cali_search_row = 13
									PF8
								END IF
								CALL EMReadScreen(end_of_data, 11, cali_search_row, 39)
								IF UCASE(end_of_data) = "END OF DATA" THEN 
									worker_not_found = TRUE
									err_msg = err_msg & vbCr & "* Worker at position " & position_array & " was not found."
									EXIT DO
								END IF
							ELSE
								CALL EMReadScreen(full_position, 20, cali_search_row, 18)
								full_position = replace(full_position, " ", "")
								worker_array = worker_array & full_position & ","
								EXIT DO
							END IF
						LOOP
					ELSEIF len(position_array) = 11 THEN 
						worker_array = worker_array & position_array & ","
					ELSEIF len(position_array) <> 8 AND len(position_array) <> 11 THEN 
						err_msg = err_msg & vbCr & "* CALI position " & position_array & " is not valid."
					END IF
				END IF
			END IF
			
			'Modifying the date range for further use (including data validation)
			IF minimum_date_range = "30 days" THEN 
				minimum_days = 30
			ELSEIF minimum_date_range = "60 days" THEN
				minimum_days = 60
			ELSEIF minimum_date_range = "90 days" THEN 
				minimum_days = 90
			ELSEIF minimum_date_range = "120 days" THEN 
				minimum_days = 120
			ELSEIF minimum_date_range = "6 months" THEN 
				minimum_days = 180
			ELSEIF minimum_date_range = "1 year" THEN 
				minimum_days = 365
			ELSEIF minimum_date_range = "2 years" THEN 
				minimum_days = 730
			ELSEIF minimum_date_range = "No payment" THEN 
				minimum_days = 0
			END IF
			
			IF maximum_date_range = "60 days" THEN 
				max_days = 60
			ELSEIF maximum_date_range = "90 days" THEN 
				max_days = 90
			ELSEIF maximum_date_range = "120 days" THEN 
				max_days = 120
			ELSEIF maximum_date_range = "6 months" THEN 
				max_days = 180
			ELSEIF maximum_date_range = "1 year" THEN 
				max_days = 365
			ELSEIF maximum_date_range = "2 years" THEN 
				max_days = 730
			ELSEIF maximum_date_range = "Open Ended" THEN 
				max_days = 100000   'otherwise known as "a bunch"
			END IF
			
			IF max_days = minimum_days 															THEN err_msg = err_msg & vbCr & "* Please select maximum number of days that differs from the minimum number of days."
			IF minimum_date_range = "Select one..." 											THEN err_msg = err_msg & vbCr & "* Please select a minimum date range."
			IF maximum_date_range = "Select one..." AND minimum_date_range <> "No payment" 		THEN err_msg = err_msg & vbCr & "* Please select a maximum date range."
			IF position = "" AND ButtonPressed = Individual_Run_button							THEN err_msg = err_msg & vbCr & "* Please select a worker."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
			
	LOOP UNTIL err_msg = ""
	PF3
	'Confirming the date range
	date_confirmation_msg = MsgBox("Last Payment Date Range" & vbCr & "=====================" & vbCr & DateAdd("D", -(max_days), date) & " - " & DateAdd("D", -(minimum_days), date) & vbCr & vbCr & "Press YES to confirm." & vbCr & "Press NO to change the date range." & vbCr & "Press CANCEL to stop the script.", vbYesNoCancel)
	IF date_confirmation_msg = vbCancel THEN stopscript
	
LOOP UNTIL date_confirmation_msg = vbYes
				
CALL check_for_PRISM (False) 'Check to see if PRISM is locked

worker_array = left(worker_array, len(worker_array) - 1)
worker_array = split(worker_array)

FOR EACH prism_worker IN worker_array
	IF prism_worker <> "" THEN 
		'Opening the Excel file
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = True
		Set objWorkbook = objExcel.Workbooks.Add() 
		objExcel.DisplayAlerts = True
		
		objExcel.Cells(1, 1).Value = "CASE NUMBER"  'Creates headings for the Excel file
		objExcel.Cells(1, 2).Value = "LAST PAY DATE"
		objExcel.Cells(1, 3).Value = "NCP NAME"
		objExcel.Cells(1, 4).Value = "CP NAME"
		
		excel_row = 2
	
		team_dropdown = left(right(prism_worker, 5), 3)
		position = right(prism_worker, 2)
	
		CALL refresh_CALI_screen
		CALL navigate_to_PRISM_screen("CALI")  'Navigate to CALI, remove any case number entered, and set the team and position
		
		EMWriteScreen county_cali_code, 20, 18
		EMWriteScreen "001", 20, 30
		EMWriteScreen "            ", 20, 58
		EMWriteScreen "  ", 20, 69
		EMWriteScreen team_dropdown, 20, 40
		EMWriteScreen position, 20, 49
		transmit
	
		EMReadScreen error_msg, 20, 24, 2   'If there is an error message because the user selected a team and position that does
		error_msg = trim (error_msg)        'not have a caseload, end script with an error message.
		IF error_msg <> "" THEN 
			MsgBox "You have selected an invalid caseload."
		ELSE
			CALL find_variable("Worker Id: ", worker_id, 9) 'Find worker id
			
			cali_row = 8  'navigates to the first case listed in CALI 
			DO
				EMReadScreen end_of_data, 11, cali_row, 32    'Goes through all the cases from CALI and puts the case number in the array, separated by "," as a delimeter.
				EMReadScreen PRISM_case_number, 14, cali_row, 7
				IF end_of_data <> "End of Data" THEN
					PRISM_case_number = replace(PRISM_case_number, "  ", "-")
					arrears_array = arrears_array & PRISM_case_number & ","
					cali_row = cali_row + 1
					excel_row = excel_row + 1
				END IF
				IF cali_row = 19 THEN    'Navigate to a new page 
					cali_row = 8
					PF8
				END IF
			LOOP UNTIL end_of_data = "End of Data"
			
			arrears_array = trim(arrears_array)  'Removes excess spaces from the arrray
			arrears_array = split(arrears_array, ",") 'Creates an array by recongizing the "," as a delimeter 
				
			excel_row = 2
			total_cases = 0
			total_paying_cases = 0
			target_cases = 0
					
			FOR EACH PRISM_case_number IN arrears_array
				CALL navigate_to_PRISM_screen("PALC") 'Navigate to the PALC screen
				IF PRISM_case_number <> "" THEN 
					'objExcel.Cells(excel_row, 1).Value = PRISM_case_number
					case_prefix = left(PRISM_case_number, 10)  'Format the case number
					case_suffix = right(PRISM_case_number, 2)
			
					EMWriteScreen case_prefix, 20, 9 'Write case number in PALC screen
					EMWriteScreen case_suffix, 20, 20
					EMWriteScreen date, 20, 49
					transmit
					
					EMReadScreen cp_name, 30, 4, 12			
					EMReadScreen ncp_name, 30, 5, 12
					EMReadScreen access_denied, 40, 24, 2
					IF InStr(UCASE(access_denied), "ACCESS DENIED") <> 0 THEN 
						ncp_name = "ACCESS DENIED"
						cp_name = "ACCESS DENIED"
					END IF
		
					EMReadScreen end_of_data, 11, 9, 32
					IF end_of_data <> "End of Data" THEN      
						EMReadScreen last_pay_date, 6, 9, 7 'Reading the last payment date
						'For run mode -- looking for all cases that have never received a payment.
						IF minimum_date_range = "No payment" THEN 
							last_pay_date = trim(last_pay_date)
							'If a payment has never been received...
							IF last_pay_date = "" THEN
								'...if we are only looking for arrears-only cases...
								IF arrears_only_checkbox = 1 THEN 
									'...go to CAST
									CALL navigate_to_PRISM_screen("CAST")
									'...reading the arrears-only field
									EMReadScreen arrears_only_case, 1, 12, 77
									IF arrears_only_case = "Y" THEN 
										'...writing the case information when the case has never received a payment and is arrears-only and that's what we're looking for.
										objExcel.Cells(excel_row, 1).Value = PRISM_case_number
										IF InStr(UCASE(access_denied), "ACCESS DENIED") <> 0 THEN 
											objExcel.Cells(excel_row, 2).Value = "NOT AVAILABLE"
										ELSE
											objExcel.Cells(excel_row, 2).Value = "NO PAYMENT"
										END IF
										objExcel.Cells(excel_row, 3).Value = ncp_name
										objExcel.Cells(excel_row, 4).Value = cp_name
										excel_row = excel_row + 1
									END IF
								
								'...if we are NOT looking for arrears-only cases...
								ELSE
									'...write the case information...
									objExcel.Cells(excel_row, 1).Value = PRISM_case_number
									IF InStr(UCASE(access_denied), "ACCESS DENIED") <> 0 THEN 
										objExcel.Cells(excel_row, 2).Value = "NOT AVAILABLE"
									ELSE
										objExcel.Cells(excel_row, 2).Value = "NO PAYMENT"
									END IF
									objExcel.Cells(excel_row, 3).Value = ncp_name
									objExcel.Cells(excel_row, 4).Value = cp_name
									excel_row = excel_row + 1
								END IF
								
							'...if a payment has been received on this case...
							ELSE
								'...if we are looking at arrears-only cases...
								IF arrears_only_checkbox = 1 THEN 
									CALL navigate_to_PRISM_screen("CAST")
									EMReadScreen arrears_only_case, 1, 12, 77
									IF arrears_only_case = "Y" THEN total_paying_cases = total_paying_cases + 1
									
								'...if we are not looking for arrears-only cases, then increase the total_paying_cases
								ELSE
									total_paying_cases = total_paying_cases + 1
								END IF
							END IF
						
						'...if we are looking at a specified date range...
						ELSE
							pay_date = left(right(last_pay_date, 4), 2) & "/" & right(last_pay_date, 2) & "/" & left(last_pay_date, 2)  'Change format of the last payment date to MM/DD/YY instead of YYMMDD.
							pay_date = CDATE(pay_date) 'CDate will allow script to recognize the string as a date. 
							IF datediff("D", pay_date, date) > minimum_days AND DateDiff("D", pay_date, date) < max_days THEN 'Checks to see how many days have elapsed since last payment.  If there has not been a payment in the past 30 days, prints a 
								IF arrears_only_checkbox = 1 THEN 
									CALL navigate_to_PRISM_screen("CAST")
									EMReadScreen arrears_only_case, 1, 12, 77
									IF arrears_only_case = "Y" THEN 
										objExcel.Cells(excel_row, 1).Value = PRISM_case_number
										objExcel.Cells(excel_row, 2).Value = pay_date
										IF InStr(UCASE(access_denied), "ACCESS DENIED") <> 0 THEN objExcel.Cells(excel_row, 2).Value = "NOT AVAILABLE"
										objExcel.Cells(excel_row, 3).Value = ncp_name
										objExcel.Cells(excel_row, 4).Value = cp_name
										excel_row = excel_row + 1							
									END IF
								ELSE
									objExcel.Cells(excel_row, 1).Value = PRISM_case_number
									objExcel.Cells(excel_row, 2).Value = pay_date
									IF InStr(UCASE(access_denied), "ACCESS DENIED") <> 0 THEN objExcel.Cells(excel_row, 2).Value = "NOT AVAILABLE"
									objExcel.Cells(excel_row, 3).Value = ncp_name
									objExcel.Cells(excel_row, 4).Value = cp_name
									excel_row = excel_row + 1
								END IF
							ELSE
								IF arrears_only_checkbox = 1 THEN 
									CALL navigate_to_PRISM_screen("CAST")
									EMReadScreen arrears_only_case, 1, 12, 77
									IF arrears_only_case = "Y" THEN total_paying_cases = total_paying_cases + 1
								ELSE
									total_paying_cases = total_paying_cases + 1
								END IF
							END IF
						END IF
					ELSE
						objExcel.Cells(excel_row, 1).Value = case_prefix & "-" & case_suffix
						objExcel.Cells(excel_row, 3).Value = ncp_name
						objExcel.Cells(excel_row, 4).Value = cp_name
						excel_row = excel_row + 1
					END IF
					total_cases = total_cases + 1 ' Add the case to the total cases tally.
				END IF
			NEXT
			
			collection_rate = total_paying_cases / total_cases  'calculate collection rate
			three_percent_cases = total_cases * .03 
			
			objExcel.Columns(1).ColumnWidth = 18 'widen columns
			objExcel.Columns(2).ColumnWidth = 14
			objExcel.Columns(3).ColumnWidth = 30
			objExcel.Columns(4).ColumnWidth = 30
			excel_row = excel_row + 1 'populate statistics for the report
			objExcel.Cells(excel_row, 1).Value = "Stats for " & worker_id
			excel_row = excel_row + 1
			objExcel.Cells(excel_row, 1).Value = "Total number of paying cases:"
			objExcel.Cells(excel_row, 3).Value = CStr(total_paying_cases)
			excel_row = excel_row +1
			objExcel.Cells(excel_row, 1).Value = "Total number of cases on caseload:"
			objExcel.Cells(excel_row, 3).Value = CStr(total_cases)
			excel_row = excel_row + 1
			objExcel.Cells(excel_row, 1).Value = "Collection rate: "
			objExcel.Cells(excel_row, 3).Value = (collection_rate * 100) & "%" 
			excel_row = excel_row + 1
			objExcel.Cells(excel_row, 1).Value = "Number of cases to change by 3%: "
			objExcel.Cells(excel_row, 3).Value = CStr(Round(three_percent_cases + .5))
			excel_row = excel_row + 1
			objExcel.Cells(excel_row, 1).Value = "Report run:"
			objExcel.Cells(excel_row, 3).Value = date & " at " & time
		END IF
	END IF
NEXT


script_end_procedure("Success!!")
