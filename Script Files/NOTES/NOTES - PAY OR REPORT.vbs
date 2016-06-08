'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - PAY OR REPORT.vbs"
start_time = timer
'
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

FUNCTION convert_month_text_to_number(month_text, month_number)
	IF month_text = "January" THEN 
		month_number = 1
	ELSEIF month_text = "February" THEN 
		month_number = 2
	ELSEIF month_text = "March" THEN 
		month_number = 3
	ELSEIF month_text = "April" THEN 
		month_number = 4 
	ELSEIF month_text = "May" THEN 
		month_number = 5 
	ELSEIF month_text = "June" THEN 
		month_number = 6
	ELSEIF month_text = "July" THEN
		month_number = 7
	ELSEIF month_text = "August" THEN 
		month_number = 8
	ELSEIF month_text = "September" THEN 
		month_number = 9 
	ELSEIF month_text = "October" THEN 
		month_number = 10
	ELSEIF month_text = "November" THEN 
		month_number = 11
	ELSEIF month_text = "December" THEN 
		month_number = 12
	END IF
END FUNCTION

FUNCTION find_second_friday(first_of_month, date_to_pay)
	pay_month = DatePart("M", first_of_month)
	pay_year = DatePart("YYYY", first_of_month)
	num_days = 1
	fridays = 0
	DO
		date_to_pay = pay_month & "/" & num_days & "/" & pay_year
		IF weekday(date_to_pay) = 6 THEN fridays = fridays + 1
		num_days = num_days + 1
	LOOP UNTIL weekday(date_to_pay) = 6 AND fridays = 2
		
	date_to_pay = DateAdd("D", 0, date_to_pay)
END FUNCTION

FUNCTION create_pay_or_report_dlg(num_of_months, pay_or_report_dates_array)

	BeginDialog pay_or_report_dialog, 0, 0, 291, (105 + (num_of_months * 20)), "Pay or Report"
	EditBox 50, 10, 55, 15, Order_date
	ComboBox 175, 11, 110, 15, ""+chr(9)+CAO_array, CAO_list
	EditBox 70, 40, 165, 15, purge_condition
	Text 10, 15, 40, 10, "Order date:"
	Text 120, 15, 55, 10, "County Attorney"
	Text 5, 45, 60, 10, "Purge Condition:"
	
	'Based on the number of months, the script will dynamically build the size of the dialog and populate the editboxes.
	IF num_of_months >= 1 THEN 
		EditBox 80, 70, 50, 15, pay_or_report_dates_array(0, 0)
		EditBox 190, 70, 50, 15, pay_or_report_dates_array(0, 1)
		Text 20, 75, 60, 10, "Payment Due:"
		Text 140, 75, 45, 10, "Report Date:"
	END IF
	IF num_of_months >= 2 THEN 
		EditBox 80, 90, 50, 15, pay_or_report_dates_array(1, 0)
		EditBox 190, 90, 50, 15, pay_or_report_dates_array(1, 1)
		Text 20, 95, 60, 10, "Payment Due:"
		Text 140, 95, 45, 10, "Report Date:"
	END IF
	IF num_of_months >= 3 THEN 
		EditBox 80, 110, 50, 15, pay_or_report_dates_array(2, 0)
		EditBox 190, 110, 50, 15, pay_or_report_dates_array(2, 1)
		Text 20, 115, 60, 10, "Payment Due:"
		Text 140, 115, 45, 10, "Report Date:"
	END IF
	IF num_of_months >= 4 THEN 
		EditBox 80, 130, 50, 15, pay_or_report_dates_array(3, 0)
		EditBox 190, 130, 50, 15, pay_or_report_dates_array(3, 1)
		Text 20, 135, 60, 10, "Payment Due:"
		Text 140, 135, 45, 10, "Report Date:"
	END IF	
	IF num_of_months >= 5 THEN 
		EditBox 80, 150, 50, 15, pay_or_report_dates_array(4, 0)
		EditBox 190, 150, 50, 15, pay_or_report_dates_array(4, 1)
		Text 20, 155, 60, 10, "Payment Due:"
		Text 140, 155, 45, 10, "Report Date:"
	END IF	
	IF num_of_months = 6 THEN 
		EditBox 80, 170, 50, 15, pay_or_report_dates_array(5, 0)
		EditBox 190, 170, 50, 15, pay_or_report_dates_array(5, 1)
		Text 20, 175, 60, 10, "Payment Due:"
		Text 140, 175, 45, 10, "Report Date:"
	END IF	
	
	Text 15, 90 + (20 * num_of_months), 65, 10, "Worker Signature"
	EditBox 85, 85 + (20 * num_of_months), 55, 15, worker_signature
	ButtonGroup ButtonPressed
		OkButton 165, 85 + (20 * num_of_months), 50, 15
		CancelButton 225, 85 + (20 * num_of_months), 50, 15
	EndDialog
	
	DO
		err_msg = ""
		Dialog pay_or_report_dialog
			IF ButtonPressed = 0 THEN stopscript
			IF order_date = "" THEN err_msg = err_msg & vbCr & "* Please enter an Order Date."
			IF CAO_list = "" THEN err_msg = err_msg & vbCr & "* Please select a County Attorney."
			IF purge_condition = "" THEN err_msg = err_msg & vbCr & "* Please enter a Purge Condition."
			FOR a = 0 to (num_of_months - 1)
				FOR b = 0 to 1
					IF IsDate(pay_or_report_dates_array(a, b)) = False AND b = 0 THEN 
						err_msg = err_msg & vbCr & "* Pay Date " & (a + 1) & " is not formatted as a date."
					ELSEIF IsDate(pay_or_report_dates_array(a, b)) = False AND b = 1 THEN 
						err_msg = err_msg & vbCr & "* Report Date " & (a + 1) & " is not formatted as a date."
					ELSEIF pay_or_report_dates_array(a, b) = "" AND b = 0 THEN 
						err_msg = err_msg & vbCr & "* Please enter a valid date for Pay Date " & (a + 1) & "."
					ELSEIF pay_or_report_dates_array(a, b) = "" AND b = 1 THEN 
						err_msg = err_msg & vbCr & "* Please enter a valid date for Report Date " & (a + 1) & "."
					END IF
				NEXT
			NEXT
			IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your CAAD note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve to condition."
	LOOP UNTIL err_msg = ""
END FUNCTION 

first_year = Cstr(DatePart("YYYY", date))
second_year = Cstr(DatePart("YYYY", DateAdd("YYYY", 1, date)))
third_year = Cstr(DatePart("YYYY", DateAdd("YYYY", 2, date)))

'=====THE DIALOGS=====
BeginDialog case_number_dialog, 0, 0, 176, 95, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  DropListBox 55, 25, 45, 10, "Month..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", month_list
  DropListBox 105, 25, 35, 15, "Year..."+chr(9)+first_year+chr(9)+second_year+chr(9)+third_year, year_list
  DropListBox 75, 45, 45, 15, "# of Months..."+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6", num_of_months
  ButtonGroup ButtonPressed
    OkButton 70, 75, 50, 15
    CancelButton 120, 75, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 45, 10, "First month:"
  Text 5, 50, 65, 10, "Number of Months"
EndDialog


'The Script
'Connecting to BlueZone
EMConnect ""

'Case number display dialog
DO
	err_msg = ""
	Dialog case_number_dialog
	If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then err_msg = err_msg & vbCr & "* Please enter your PRISM case number in a valid format: ''XXXXXXXXXX-XX''"
		IF month_list = "Month..." THEN err_msg = err_msg & vbCr & "* Please select a month."
		IF year_list = "Year..." THEN err_msg = err_msg & vbCr & "* Please select a year."
		IF num_of_months = "# of Months..." THEN err_msg = err_msg & vbCr & "* Please select the number of months."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

ReDim pay_or_report_dates_array(num_of_months, 1)

'Creating first of months and second Fridays
first_month = (month_list & "/01/" & year_list)
	
FOR i = 0 to (num_of_months - 1)
	pay_or_report_dates_array(i, 0) = DateAdd("M", i, first_month)
	CALL find_second_friday(pay_or_report_dates_array(i, 0), pay_or_report_dates_array(i, 1))
	pay_or_report_dates_array(i, 0) = CStr(pay_or_report_dates_array(i, 0))
	pay_or_report_dates_array(i, 1) = CStr(pay_or_report_dates_array(i, 1))
NEXT
	
CALL convert_array_to_droplist_items(county_attorney_array, CAO_array)

CALL create_pay_or_report_dlg(num_of_months, pay_or_report_dates_array)

CALL check_for_PRISM(False)

'Going to CAWT screen
call navigate_to_PRISM_screen("CAWT")
FOR k = 0 to (num_of_months - 1)
	FOR j = 0 to 1
		PF5									'adding a note
		EMWriteScreen "FREE", 4,37					'adding a worklist
		IF j = 0 THEN 								'array argument 0 is the pay date (first of the month)
			EMWriteScreen "Check for purge payments, due today.", 10, 4		'adding a line in the worklist
		ELSEIF j = 1 THEN 							'array argument 1 is the report date (second Friday of the month)
			EMWriteScreen "Check for purge payments, report date.", 10, 4
		END IF
		CALL create_mainframe_friendly_date(pay_or_report_dates_array(k, j), 17, 21, "YYYY")		'creating the worklists in PRISM
		transmit								'adding the worklist to CAWT
		PF3									'backing out of worklist
	NEXT
NEXT


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

'Entering case number
EMWriteScreen case_number, 20, 8
'Add a new CAAD note
PF5
'CAAD type
EMWriteScreen "FREE", 4, 54
'The CAAD Note
EMSetCursor 16, 4					'Because the cursor does not default to this location
call write_new_line_in_PRISM_case_note("Pay or Report Information")
call write_editbox_in_PRISM_case_note("Purge Condition", purge_condition, 6)  
call write_editbox_in_PRISM_case_note("Order Date", Order_date, 6)
call write_editbox_in_PRISM_case_note("County Attorney", CAO_list, 6)
call write_new_line_in_PRISM_case_note("---")	
call write_new_line_in_PRISM_case_note(worker_signature)

script_end_procedure("Success!!")
