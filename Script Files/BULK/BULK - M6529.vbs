'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "BULK - M6529.vbs" 
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

'This function takes a date and addes leading zeros if necessary to format MM/DD/YYYY
FUNCTION change_date_format(date_to_format)
		month3 = DatePart("M", date_to_format)
		day3 = DatePart("D", date_to_format)
		year3 = DatePart("YYYY", date_to_format)
		if len(month3) = 1 then 
			month3_s = "0" & month3
		else
			month3_s = month3
		end if
		if len(day3) = 1 then 
			day3_s = "0" & day3
		else
			day3_s = day3
		end if
		date_to_format = month3_s & "/" & day3_s & "/" & year3
		change_date_format = date_to_format
END FUNCTION

' >>>>> DETERMINING FIRST DAY OF THE MONTH <<<<< 
current_month = DatePart("M", date)
if len(current_month) = 1 then current_month_s = "0" & current_month
current_year = DatePart("YYYY", date)
current_date = current_month_s & "/01/" & current_year


' >>>>> DETERMINING LAST DAY OF THE MONTH <<<<< 
next_month = DateAdd("M", 1, current_date)
next_month_minus1 = DateAdd("D", -1, next_month)
next_month_minus1_s = change_date_format(next_month_minus1)

'>>>>>> DETERMINING A DATE 3 MONTHS AGO  <<<<<<	
current_month_minus3 = DateAdd("M", -3, date) 'variable for the current date minus three - this returns a date format
current_month_minus3_s = change_date_format(current_month_minus3)

' >>>>> THE SCRIPT <<<<<
EMConnect ""

'>>>>> GOING TO USWT <<<<<
Call navigate_to_Prism_screen("USWT")

' >>>>> SELECTING THE SPECIFIC WORKLIST INFO <<<<<
EMWriteScreen "M6529", 20, 30
EMWriteScreen current_date, 20, 48   'Select worklists that are due this calendar month 
EMWriteScreen next_month_minus1_s, 20, 63
transmit

USWT_row = 7
COUNT = 0
SCROLL = 0
' >>>>> STARTING THE DO LOOP. THE SCRIPT NEEDS TO HANDLE THESE CASES ONE AT A TIME <<<<<
DO
	EMReadScreen USWT_type, 5, USWT_row, 45
	EMReadScreen USWT_date, 8, USWT_row, 73
	IF USWT_type = "M6529" THEN
		
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		EMWriteScreen "D", USWT_row, 4
		transmit
		
		purge = false 'Reset the purge variable
		Call navigate_to_Prism_screen("CAFS")
		EMReadScreen monthly_accrual, 13, 9, 26
		EMReadScreen monthly_nonaccrual, 13, 10, 26
		total_due = (ccur(monthly_accrual) + ccur(monthly_nonaccrual))*1.2*3
		

		Call navigate_to_PRISM_screen ("PALC")

		
		EMWriteScreen current_month_minus3_s, 20, 35  'Checks for payments for the last three months
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
		If trim(case_alloc_amt) = "" then case_alloc_amt = 0												'Reading the amt allocated
		If proc_type = "FTS" or proc_type = "MCE" or proc_type = "NOC" or proc_type = "IFC" or proc_type = "OST" or _	
		proc_type = "PCA" or proc_type = "PIF" or proc_type = "STJ" or proc_type = "STS" or proc_type = "FTJ" then 		'If proc type is one of these, it's involuntary. Else, it's voluntary.
			total_involuntary_alloc = total_involuntary_alloc + abs(case_alloc_amt)							'Adds the alloc amt for involuntary
		Else
			total_voluntary_alloc = total_voluntary_alloc + abs(case_alloc_amt)							'Adds the alloc amt for voluntary
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
msgbox_text = "---PAYMENT BREAKDOWN FOR " & current_month_minus3 & " THROUGH " & date & "---" & chr(10) & chr(10)_
			& "Involuntary: " & formatCurrency(ccur(total_involuntary_alloc)) & chr(10) & "Voluntary: " & formatCurrency(ccur(total_voluntary_alloc)) _
			& chr(10) & chr(10) & "Monthly accrual: " & formatCurrency(ccur(monthly_accrual)) & chr(10) _
			& "Monthly Non-accrual: " & formatCurrency(ccur(monthly_nonaccrual)) & chr(10) & chr(10) & chr(10) _ 
			& "Total due: " & formatCurrency(ccur(total_due)) &chr(10)& "Continue Interest Suspension?"

continue = msgbox (msgbox_text, 4, "Continue Interest Suspension?")

IF continue = 6 then 
'msgbox "User selected to continue the interest suspension."
purge = true 
count = count + 1
END IF	
	IF continue = 7 then 'User selected not to add the employer
'msgbox "User selected to end the interest suspension."
count = count + 1
END IF

	
		Call navigate_to_PRISM_screen ("CAWT")
		EMWriteScreen "M6529", 20, 29
		EMWriteScreen current_date, 20, 47
		EMWriteScreen next_month_minus1, 20, 62
		EMWriteScreen USWT_case_number, 20, 8
		transmit

		' >>>>> IF THE WORKLIST ITEM IS ELIGIBLE TO BE PURGED, THE SCRIPT PURGES...
		IF purge = True THEN
 			Call navigate_to_PRISM_screen ("CAWT")
			EMWriteScreen "M6529", 20, 29
			EMWriteScreen current_date, 20, 47
			EMWriteScreen next_month_minus1, 20, 62
			EMWriteScreen USWT_case_number, 20, 8
			transmit
			EMReadScreen CAWT_type, 5, 8, 8
			IF CAWT_type = "M6529" then
				EMWriteScreen "P", 8, 4
				transmit
			END IF
			Call navigate_to_PRISM_screen ("CAAD")
			PF5
			EMWriteScreen "A", 3, 29
			EMWriteScreen "M6529", 4, 54
			EMSetCursor 16, 4
caad_text = "---PAYMENT BREAKDOWN FOR " & current_month_minus3 & " THROUGH " & date & "---  "_
			& "Involuntary: " & formatCurrency(ccur(total_involuntary_alloc)) & " " & "Voluntary: " & formatCurrency(ccur(total_voluntary_alloc)) _
			& "  " & "Monthly accrual: " & formatCurrency(ccur(monthly_accrual)) & " " _
			& "Monthly Non-accrual: " & formatCurrency(ccur(monthly_nonaccrual)) & "     " _ 
			& "Total due: " & formatCurrency(ccur(total_due)) & " "
			write_variable_in_CAAD(caad_text) 
			transmit				
		END IF

total_voluntary_alloc = "0"
total_involuntary_alloc = "0"


		'  ...  IF THE WORKLIST ITEM IS NOT ELIGIBLE TO BE PURGED, THE SCRIPT INCREASES USWT_ROW + 1 <<<<<
			Call navigate_to_PRISM_screen ("USWT")

			EMWriteScreen "M6529", 20, 30
			EMWriteScreen current_date, 20, 48
			EMWriteScreen next_month_minus1, 20, 63
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
		
	End If
LOOP UNTIL USWT_type <> "M6529"

script_end_procedure("Success!  " & Count & " worklists have been processed.")

