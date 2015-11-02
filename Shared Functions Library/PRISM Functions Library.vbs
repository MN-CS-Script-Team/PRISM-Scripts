'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script contains functions that the other BlueZone scripts use very commonly. The
'other BlueZone scripts contain a few lines of code that run this script and get the 
'functions. This saves me time in writing and copy/pasting the same functions in
'many different places. Only add functions to this script if they've been tested by
'the workgroups. This document is actively used by live scripts, so it needs to be
'functionally complete at all times.
'
'Here's the code to add, including stats gathering pieces (without comments of course):
'
''LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
'Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
'If beta_agency = "" then 			'For scriptwriters only
'	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
'ElseIf beta_agency = True then		'For beta agencies and testers
'	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
'Else								'For most users
'	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
'End if
'Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
'req.open "GET", url, False									'Attempts to open the URL
'req.send													'Sends request
'If req.Status = 200 Then									'200 means great success
'	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'	Execute req.responseText								'Executes the script code
'ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
'	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
'			vbCr & _
'			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
'			vbCr & _
'			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
'			vbTab & "- The name of the script you are running." & vbCr &_
'			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
'			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
'			vbTab & vbTab & "responsible for network issues." & vbCr &_
'			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
'			vbCr & _
'			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
'			vbCr &_
'			"URL: " & url
'			StopScript
'END IF

'GLOBAL CONSTANTS----------------------------------------------------------------------------------------------------
'Declares variables (thinking of option explicit in the future)
Dim checked, unchecked, cancel, OK

checked = 1			'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0			'Value for cancel button in dialogs
OK = -1				'Value for OK button in dialogs

'SHARED FUNCTIONS----------------------------------------------------------------------------------------------------

Function attn
  EMSendKey "<attn>"
  EMWaitReady -1, 0
End function

function back_to_SELF
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

Function check_for_PRISM(end_script)
	PF11
	PF10
	CALL find_variable("PLEASE ENTER YOUR ", timed_out, 8)
	IF timed_out = "PASSWORD" THEN 
		IF end_script = True THEN 
			If PRISM_check <> "PRISM" then script_end_procedure("You do not appear to be in PRISM. You may be passworded out. Please check your PRISM screen and try again.")
		ELSE
			If PRISM_check <> "PRISM" then MsgBox "You do not appear to be in PRISM. You may be passworded out."
		END IF
	END IF
END FUNCTION

Function convert_array_to_droplist_items(array_to_convert, output_droplist_box)
	For each item in array_to_convert
		If output_droplist_box = "" then 
			output_droplist_box = item
		Else
			output_droplist_box = output_droplist_box & chr(9) & item
		End if
	Next
End Function

FUNCTION create_mainframe_friendly_date(date_variable, screen_row, screen_col, year_type) 
	var_month = datepart("m", date_variable)
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month & "/", screen_row, screen_col
	var_day = datepart("d", date_variable)
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day & "/", screen_row, screen_col + 3
	If year_type = "YY" then
		var_year = right(datepart("yyyy", date_variable), 2)
	ElseIf year_type = "YYYY" then
		var_year = datepart("yyyy", date_variable)
	Else
		MsgBox "Year type entered incorrectly. Fourth parameter of function create_mainframe_friendly_date should read ""YYYY"" or ""YY"". The script will now stop."
		StopScript
	END IF
	EMWriteScreen var_year, screen_row, screen_col + 6
END FUNCTION

FUNCTION date_converter_PALC_PAPL (date_variable)

	date_year = left (date_variable, 2)
	date_day = right (date_variable, 2)
	date_month = right (left (date_variable, 4), 2)
	
	date_variable = date_month & "/" & date_day & "/" & date_year 
END FUNCTION

FUNCTION change_date_format_from_MDYYYY_to_MMDDYYY(date_to_format)
		month3 = DatePart("M", date_to_format)
		day3 = DatePart("D", date_to_format)
		year3 = DatePart("YYYY", date_to_format)
		if len(month3) = 1 then month3 = 0 & month3
		if len(day3) = 1 then day3 = 0 & day3
		date_to_format = month3 & "/" & day3 & "/" & year3
		change_date_format = date_to_format
END FUNCTION

Function end_excel_and_script
  objExcel.Workbooks.Close
  objExcel.quit
  stopscript
End function

Function enter_PRISM_case_number(case_number_variable, row, col)
	EMSetCursor row, col
	EMSendKey replace(case_number_variable, "-", "")                                                                                                                                       'Entering the specific case indicated
	EMSendKey "<enter>"
	EMWaitReady 0, 0
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

'This function fixes the case for a phrase. For example, "ROBERT P. ROBERTSON" becomes "Robert P. Robertson". 
'	It capitalizes the first letter of each word.
Function fix_case(phrase_to_split, smallest_length_to_skip)									'Ex: fix_case(client_name, 3), where 3 means skip words that are 3 characters or shorter
	phrase_to_split = split(phrase_to_split)											'splits phrase into an array
	For each word in phrase_to_split												'processes each word independently
		If word <> "" then													'Skip blanks
			first_character = ucase(left(word, 1))									'grabbing the first character of the string, making uppercase and adding to variable
			remaining_characters = LCase(right(word, len(word) -1))						'grabbing the remaining characters of the string, making lowercase and adding to variable
			If len(word) > smallest_length_to_skip then								'skip any strings shorter than the smallest_length_to_skip variable
				output_phrase = output_phrase & first_character & remaining_characters & " "		'output_phrase is the output of the function, this combines the first_character and remaining_characters
			Else															
				output_phrase = output_phrase & word & " "							'just pops the whole word in if it's shorter than the smallest_length_to_skip variable
			End if
		End if
	Next
	phrase_to_split = output_phrase												'making the phrase_to_split equal to the output, so that it can be used by the rest of the script.
End function

function navigate_to_MAXIS_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(x) and STAT_note_check <> "NOTE" then 
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen y, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen x, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen y, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    End if
  End if
End function

Function navigate_to_PRISM_screen(x) 'x is the name of the screen
  EMWriteScreen x, 21, 18
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

Function PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
End function

Function PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

Function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

Function PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
End function

Function PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
End function

Function PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
End function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

Function PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
End function

Function PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
End function

Function PF13
  EMSendKey "<PF13>"
  EMWaitReady 0, 0
End function

Function PF14
  EMSendKey "<PF14>"
  EMWaitReady 0, 0
End function

Function PF15
  EMSendKey "<PF15>"
  EMWaitReady 0, 0
End function

Function PF16
  EMSendKey "<PF16>"
  EMWaitReady 0, 0
End function

Function PF17
  EMSendKey "<PF17>"
  EMWaitReady 0, 0
End function

Function PF18
  EMSendKey "<PF18>"
  EMWaitReady 0, 0
End function

Function PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
End function

function PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

Function PF21
  EMSendKey "<PF21>"
  EMWaitReady 0, 0
End function

Function PF22
  EMSendKey "<PF22>"
  EMWaitReady 0, 0
End function

Function PF23
  EMSendKey "<PF23>"
  EMWaitReady 0, 0
End function

Function PF24
  EMSendKey "<PF24>"
  EMWaitReady 0, 0
End function

Function PRISM_case_number_finder(variable_for_PRISM_case_number)
	'Searches for the case number.
	PRISM_row = 1
	PRISM_col = 1
	EMSearch "Case: ", PRISM_row, PRISM_col
	If PRISM_row <> 0 then
		EMReadScreen variable_for_PRISM_case_number, 13, PRISM_row, PRISM_col + 6
		variable_for_PRISM_case_number = replace(variable_for_PRISM_case_number, " ", "-")
	Else	'Searches again if not found, this time for "Case/Person"
		PRISM_row = 1
		PRISM_col = 1
		EMSearch "Case/Person: ", PRISM_row, PRISM_col
		If PRISM_row <> 0 then
			EMReadScreen variable_for_PRISM_case_number, 13, PRISM_row, PRISM_col + 13
			variable_for_PRISM_case_number = replace(variable_for_PRISM_case_number, " ", "-")
		End if
	End if	
	If isnumeric(left(variable_for_PRISM_case_number, 10)) = False or isnumeric(right(variable_for_PRISM_case_number, 2)) = False then variable_for_PRISM_case_number = ""
End function

Function PRISM_case_number_validation(case_number_to_validate, outcome)
  If len(case_number_to_validate) <> 13 then 
    outcome = False
  Elseif isnumeric(left(case_number_to_validate, 10)) = False then
    outcome = False
  Elseif isnumeric(right(case_number_to_validate, 2)) = False then
    outcome = False
  Elseif InStr(11, case_number_to_validate, "-") <> 11 then
    outcome = False
  Else
    outcome = True
  End if
End function

'Runs a script from GitHub.
FUNCTION run_from_GitHub(url)
	Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
	req.open "GET", url, False									'Attempts to open the URL
	req.send													'Sends request
	If req.Status = 200 Then									'200 means great success
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
				script_end_procedure("Script ended due to error connecting to GitHub.")
	END IF
END FUNCTION

function script_end_procedure(closing_message)
	If closing_message <> "" then MsgBox closing_message
	If collecting_statistics = True then
		stop_time = timer
		script_run_time = stop_time - start_time
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & stats_database_path

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
	End if
	stopscript
end function

function script_end_procedure_wsh(closing_message) 'For use when running a script outside of the BlueZone Script Host
	If closing_message <> "" then MsgBox closing_message
	If collecting_statistics = True then
		stop_time = timer
		script_run_time = stop_time - start_time
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & stats_database_path

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
	End if
	Wscript.Quit
end function

Function step_through_handling 'This function will introduce "warning screens" before each transmit, which is very helpful for testing new scripts
	'To use this function, simply replace the "Execute text_from_the_other_script" line with:
	'Execute replace(text_from_the_other_script, "EMWaitReady 0, 0", "step_through_handling")
	step_through = MsgBox("Step " & step_number & chr(13) & chr(13) & "If you see something weird on your screen (like a MAXIS or PRISM error), PRESS CANCEL then email the script writer about it. Make sure you include the step you're on.", 1)
	If step_number = "" then step_number = 1	'Declaring the variable
	If step_through = 2 then
		stopscript
	Else
		EMWaitReady 0, 0
		step_number = step_number + 1
	End if
End Function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

FUNCTION word_doc_open(doc_location, objWord, objDoc)
	'Opens Word object
	Set objWord = CreateObject("Word.Application")
	objWord.Visible = True		'We want to see it
	
	'Opens the specific Word doc
	set objDoc = objWord.Documents.Add(doc_location)
END FUNCTION

FUNCTION word_doc_update_field(field_name, variable_for_field, objDoc)
	'Simply enters the Word document field based on these three criteria
	objDoc.FormFields(field_name).Result = variable_for_field
END FUNCTION

Function write_bullet_and_variable_in_CAAD(bullet, variable)
  spaces_count = 6	'Temporary just to make it work

  EMGetCursor row, col 
  EMReadScreen line_check, 2, 15, 2
  If ((row = 20 and col + (len(bullet)) >= 78) or row = 21) and line_check = "26" then 
    MsgBox "You've run out of room in this case note. The script will now stop."
    StopScript
  End if
  If row = 21 then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSetCursor 16, 4
  End if
  variable_array = split(variable, " ")
  EMSendKey "* " & bullet & ": "
  For each variable_word in variable_array 
    EMGetCursor row, col 
    EMReadScreen line_check, 2, 15, 2
    If ((row = 20 and col + (len(variable_word)) >= 78) or row = 21) and line_check = "26" then 
      MsgBox "You've run out of room in this case note. The script will now stop."
      StopScript
    End if
    If (row = 20 and col + (len(variable_word)) >= 78) or (row = 16 and col = 4) or row = 21 then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
      EMSetCursor 16, 4
    End if
    EMGetCursor row, col 
    If (row < 20 and col + (len(variable_word)) >= 78) then EMSendKey "<newline>" & space(spaces_count) 
'    If (row = 16 and col = 4) then EMSendKey space(spaces_count)		'<<<REPLACED WITH BELOW IN ORDER TO TEST column issue
    If (col = 4) then EMSendKey space(spaces_count)
    EMSendKey variable_word & " "
    If right(variable_word, 1) = ";" then 
      EMSendKey "<backspace>" & "<backspace>" 
      EMGetCursor row, col 
      If row = 20 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSetCursor 16, 4
        EMSendKey space(spaces_count)
      Else
        EMSendKey "<newline>" & space(spaces_count)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 20 and col + (len(bullet)) >= 78) or (row = 16 and col = 4) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSetCursor 16, 4
  End if
End Function

Function write_variable_in_CAAD(variable)
  EMGetCursor row, col 
  EMReadScreen line_check, 2, 15, 2
  If ((row = 20 and col + (len(x)) >= 78) or row = 21) and line_check = "26" then 
    MsgBox "You've run out of room in this case note. The script will now stop."
    StopScript
  End if
  If (row = 20 and col + (len(x)) >= 78 + 1 ) or row = 21 then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSetCursor 16, 4
  End if
  EMSendKey variable & "<newline>"
  EMGetCursor row, col 
  If (row = 20 and col + (len(x)) >= 78) or (row = 21) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSetCursor 16, 4
  End if
End function

'-------------------------------------LOADING MAXIS FUNCTIONS BECAUSE THEY ARE MOSTLY SHARED
'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
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
			"URL: " & FuncLib_URL
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'----------------------------------------------------------------------------------------------------DEPRECIATED FUNCTIONS LEFT HERE FOR COMPATIBILITY PURPOSES
function PRISM_check_function													'DEPRECIATED 03/10/2015
	call check_for_PRISM(True)	'Defaults to True because that's how we always did it.
END function

Function write_editbox_in_PRISM_case_note(bullet, variable, spaces_count)		'DEPRECIATED 03/10/2015
	call write_bullet_and_variable_in_CAAD(bullet, variable)
End function

Function write_new_line_in_PRISM_case_note(variable)							'DEPRECIATED 03/10/2015
	call write_variable_in_CAAD(variable)
End function

FUNCTION write_value_and_transmit(input_value, PRISM_row, PRISM_col)
	EMWriteScreen input_value, PRISM_row, PRISM_col
	transmit
END FUNCTION

