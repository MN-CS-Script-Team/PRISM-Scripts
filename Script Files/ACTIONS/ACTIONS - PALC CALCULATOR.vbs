'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - PALC CALCULATOR.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'DIALOGS---------------------------------------------------------------------------
BeginDialog start_end_date_dialog, 0, 0, 171, 65, "Start and End Date Dialog"
  ButtonGroup ButtonPressed
    OkButton 120, 25, 50, 15
    CancelButton 120, 45, 50, 15
  EditBox 60, 5, 110, 15, PRISM_case_number
  EditBox 45, 25, 70, 15, start_date  'Start date for the search
  EditBox 45, 45, 70, 15, end_date  'End date for the search
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 40, 10, "Start date:"
  Text 5, 50, 40, 10, "End date:"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)

Do
	Do
		Dialog start_end_date_dialog				'Shows dialog
		If buttonpressed = 0 then stopscript		'Cancel
		call PRISM_case_number_validation(PRISM_case_number, case_number_is_valid)
		If case_number_is_valid = False then MsgBox "Your case number isn't valid. Try again."
	Loop until case_number_is_valid = True
	If isdate(start_date) = False or isdate(end_date) = False then MsgBox "You must enter valid dates for both the start and end dates."		'Because a date for both fields is required
Loop until isdate(start_date) = True and isdate(end_date) = True



'Checks to make sure PRISM isn't locked out
transmit
PRISM_check_function

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to PALC
call navigate_to_PRISM_screen("PALC")

'Entering case number and transmitting
EMSetCursor 20, 9
EMSendKey replace(PRISM_case_number, "-", "")	 	'Entering the specific case indicated

EMWriteScreen cstr(start_date), 20, 35
EMWriteScreen cstr(end_date), 20, 49
transmit								'Transmitting into it



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
		EMReadScreen case_alloc_amt, 10, row, 70													'Reading the amt allocated
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

string_for_msgbox = "---PAYMENT BREAKDOWN FOR " & start_date & " THROUGH " & end_date & "---" & chr(10) & chr(10) & "Involuntary: $" & total_involuntary_alloc & chr(10) & "Voluntary: $" & total_voluntary_alloc

MsgBox string_for_msgbox

script_end_procedure("")
