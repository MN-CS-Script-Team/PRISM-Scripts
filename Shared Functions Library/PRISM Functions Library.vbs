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
'
'
'****************************************************************************************
'*******************KEEP LISTS IN ALPHABETICAL ORDER, PLEASE!!!***************************
'****************************************************************************************
'
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
'ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Robert with details (and stops script).
'	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
'			vbCr & _
'			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
'			vbCr & _
'			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
'			vbTab & "- The name of the script you are running." & vbCr &_
'			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
'			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
'			vbTab & vbTab & "responsible for network issues." & vbCr &_
'			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
'			vbCr & _
'			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
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

FUNCTION cancel_confirmation
	If ButtonPressed = 0 then
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		If cancel_confirm = vbYes then script_end_procedure("CANCEL BUTTON SELECTED")     
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
	End if
END FUNCTION

' This is a custom function to change the format of a participant name.  The parameter is a string with the 
' client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".  
FUNCTION change_client_name_to_FML(client_name)
	client_name = trim(client_name)
	length = len(client_name)
	position = InStr(client_name, ", ")
	last_name = Left(client_name, position-1)
	first_name = Right(client_name, length-position-1)	
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
END FUNCTION

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

Function clear_line_of_text(row, start_column)
  EMSetCursor row, start_column
  EMSendKey "<EraseEof>"
  EMWaitReady 0, 0
End function


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

'This function takes in a client's name and outputs the name (accounting for hyphenated surnames) with Ucase first character
'and lcase the rest. This is like fix_case but this function is a bit more specific for names
FUNCTION fix_case_for_name(name_variable)
	name_variable = split(name_variable, " ")
	FOR EACH client_name IN name_variable
		IF client_name <> "" THEN 
			IF InStr(client_name, "-") = 0 THEN 
				client_name = UCASE(left(client_name, 1)) & LCASE(right(client_name, len(client_name) - 1))
				output_variable = output_variable & " " & client_name
			ELSE				'When the client has a hyphenated surname
				hyphen_location = InStr(client_name, "-")
				first_part = left(client_name, hyphen_location - 1)
				first_part = UCASE(left(first_part, 1)) & LCASE(right(first_part, len(first_part) - 1))
				second_part = right(client_name, len(client_name) - hyphen_location)
				second_part = UCASE(left(second_part, 1)) & LCASE(right(second_part, len(second_part) - 1))
				output_variable = output_variable & " " & first_part & "-" & second_part
			END IF
		END IF
	NEXT
	name_variable = output_variable
END FUNCTION


'This is a custom function to fix data that we are reading from PRISM that includes underscores.  The parameter is a string for the 
'variable to be searched.  The function searches the variable and removes underscores.  Then, the fix case function is called to format
'the string in the correct case.  Finally, the data is trimmed to remove any excess spaces.	
FUNCTION fix_read_data (search_string) 
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
END FUNCTION


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

function run_another_script(script_path)
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  Execute text_from_the_other_script
end function

'Runs a script from GitHub.
FUNCTION run_from_GitHub(url)
	Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
	req.open "GET", url, False									'Attempts to open the URL
	req.send													'Sends request
	If req.Status = 200 Then									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
		Execute req.responseText								'Executes the script code
	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Robert with details (and stops script).
		MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
				vbCr & _
				"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
				vbCr & _
				"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
				vbTab & "- The name of the script you are running." & vbCr &_
				vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
				vbTab & "- The name and email for an employee from your IT department," & vbCr & _
				vbTab & vbTab & "responsible for network issues." & vbCr &_
				vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
				vbCr & _
				"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
				vbCr &_
				"URL: " & url
				script_end_procedure("Script ended due to error connecting to GitHub.")
	END IF
END FUNCTION

Function save_cord_doc
  EMWriteScreen "M", 3, 29
  transmit
End function

function script_end_procedure(closing_message)
	If closing_message <> "" then MsgBox closing_message, vbInformation + vbSystemModal
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

'This code is helpful for bulk scripts. This script is used to select the caseload by the 8 digit worker ID code entered in the dialog.
FUNCTION select_cso_caseload(ButtonPressed, cso_id, cso_name)
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
			
			BeginDialog select_cso_dlg, 0, 0, 286, 145, " - Select CSO Caseload"
			EditBox 70, 55, 65, 15, cso_id
			Text 70, 80, 90, 10, cso_name
			ButtonGroup ButtonPressed
				OkButton 130, 125, 50, 15
				PushButton 180, 125, 50, 15, "UPDATE CSO", update_cso_button
				PushButton 230, 125, 50, 15, "STOP SCRIPT", stop_script_button
			Text 10, 15, 265, 30, "This script will check for worklist items coded E0014 for the following Worker ID. If you wish to change the Worker ID, enter the desired Worker ID in the box and press UPDATE CSO. When you are ready to continue, press OK."
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

'This function requires a recipient (the recipient code from the DORD screen), and the document code (also from the DORD screen).
'This function adds the document.  Some user involvement (resolving required labels, hard-copy printing) may be required.
FUNCTION send_dord_doc(recipient, dord_doc)
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen dord_doc, 6, 36
	EMWriteScreen recipient, 11, 51
	transmit
END FUNCTION
	
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
IF variable <> "" THEN  
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
END IF
End Function

Function write_variable_in_CAAD(variable)
IF variable <> "" THEN  
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
END IF
End function


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

Function write_variable_to_CORD_paragraph(variable)
	If trim(variable) <> "" THEN
		EMGetCursor noting_row, noting_col		'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 6					'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		IF noting_row < 11 THEN noting_row = 11	'Making sure it is writing in the paragraph.
		
		'Backing out of the CORD paragraph
		IF noting_row > 20 THEN 
			MsgBox "The script is attempting to write in a spot that is not supported by PRISM. Please review your CORD document for accuracy and contact a scripts administrator to have this issue resolved.", vbCritical + vbSystemModal, "Critical CORD Paragraph Error!!"
			EXIT FUNCTION
		END IF

		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")

		FOR EACH word IN variable_array

			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 75 then
				noting_row = noting_row + 1
				noting_col = 6
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)

			'Backing out of the CORD paragraph
			IF noting_row >= 20 THEN 
				MsgBox "The script is attempting to write in a spot that is not supported by PRISM. Please review your CORD document for accuracy and a scripts administrator to have this issue resolved.", vbCritical + vbSystemModal, "Critical CORD Paragraph Error!!"
				EXIT FUNCTION
			END IF
		NEXT

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 6
	End if
End function


'>>>>> CLASSES!!!!!!!!!!!!!!!!!!!!! <<<<<
'This CLASS contains properties used to populate documents
' These properties should not be used for other applications in scripts.
' Everytime you call the property, the script will try to navigate and grab the information
CLASS doc_info
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' >>>>> CP INFORMATION <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' CP name (last, first middle initial, suffix (if any))
	PUBLIC PROPERTY GET cp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_name, 50, 5, 25
		cp_name = trim(cp_name)
	END PROPERTY
	
	' CP first name
	PUBLIC PROPERTY GET cp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_first_name, 12, 8, 34
		cp_first_name = trim(replace(cp_first_name, "_", ""))
	END PROPERTY

	' CP last name
	PUBLIC PROPERTY GET cp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_last_name, 17, 8, 8
		cp_last_name = trim(replace(cp_last_name, "_", ""))
	END PROPERTY	
	
	' CP middle name
	PUBLIC PROPERTY GET cp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_middle_name, 12, 8, 56
		cp_middle_name = trim(replace(cp_middle_name, "_", ""))
	END PROPERTY
	
	' CP middle initial
	PUBLIC PROPERTY GET cp_middle_initial
		cp_middle_initial = left(cp_middle_name, 1)
	END PROPERTY
	
	' CP suffix
	PUBLIC PROPERTY GET cp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_suffix, 3, 8, 74
		cp_suffix = trim(replace(cp_suffix, "_", ""))
	END PROPERTY
	
	' CP date of birth
	PUBLIC PROPERTY GET cp_dob
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_dob, 8, 6, 24		
	END PROPERTY

	' CP social security number
	PUBLIC PROPERTY GET cp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_ssn, 11, 6, 7
	END PROPERTY
	
	' CP MCI
	PUBLIC PROPERTY GET cp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_mci, 10, 5, 7
	END PROPERTY	
	
	' CP address
	PUBLIC PROPERTY GET cp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_addr1, 30, 15, 11
		EMReadScreen cp_addr2, 30, 16, 11
		cp_addr = replace(cp_addr1, "_", "") & ", " & replace(cp_addr2, "_", "")
	END PROPERTY

	' CP address city
	PUBLIC PROPERTY GET cp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_city, 20, 17, 11
		cp_city = replace(cp_city, "_", "")
	END PROPERTY

	' CP address state
	PUBLIC PROPERTY GET cp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_state, 2, 17, 39
	END PROPERTY
	
    ' CP address zip code
	PUBLIC PROPERTY GET cp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_zip, 10, 17, 50
	END PROPERTY
	
	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>>>> NCP Information <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' NCP Name
	PUBLIC PROPERTY GET ncp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_name, 50, 5, 25
		ncp_name = trim(ncp_name)
	END PROPERTY
	
	' NCP first name
	PUBLIC PROPERTY GET ncp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_first_name, 12, 8, 34
		ncp_first_name = trim(replace(ncp_first_name, "_", ""))
	END PROPERTY

	' NCP last name
	PUBLIC PROPERTY GET ncp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_last_name, 17, 8, 8
		ncp_last_name = trim(replace(ncp_last_name, "_", ""))
	END PROPERTY	
	
	' NCP middle name
	PUBLIC PROPERTY GET ncp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_middle_name, 12, 8, 56
		ncp_middle_name = trim(replace(ncp_middle_name, "_", ""))
	END PROPERTY
	
	' NCP middle initial
	PUBLIC PROPERTY GET ncp_middle_initial
		ncp_middle_initial = left(ncp_middle_name, 1)
	END PROPERTY
	
	' NCP suffix
	PUBLIC PROPERTY GET ncp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_suffix, 3, 8, 74
		ncp_suffix = trim(replace(ncp_suffix, "_", ""))
	END PROPERTY	
	
	' NCP date of birth
	PUBLIC PROPERTY GET ncp_dob
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_dob, 8, 6, 24		
	END PROPERTY

	' NCP SSN
	PUBLIC PROPERTY GET ncp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_ssn, 11, 6, 7
	END PROPERTY
	
	' NCP MCI
	PUBLIC PROPERTY GET ncp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_mci, 10, 5, 7
	END PROPERTY	

	' NCP street address
	PUBLIC PROPERTY GET ncp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_addr1, 30, 15, 11
		EMReadScreen ncp_addr2, 30, 16, 11
		ncp_addr = replace(ncp_addr1, "_", "") & ", " & replace(ncp_addr2, "_", "")
	END PROPERTY

	' NCP address city
	PUBLIC PROPERTY GET ncp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_city, 20, 17, 11
		ncp_city = replace(ncp_city, "_", "")
	END PROPERTY

	' NCP address state
	PUBLIC PROPERTY GET ncp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_state, 2, 17, 39
	END PROPERTY
    
	' NCP address zip code
	PUBLIC PROPERTY GET ncp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_zip, 10, 17, 50
	END PROPERTY
	
	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>> Financial Information <<<
	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' monthly accrual amount
	PUBLIC PROPERTY GET monthly_accrual
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen monthly_accrual, 8, 9, 31
		monthly_accrual = trim(monthly_accrual)
	END PROPERTY
	
	' monthly non-accrual
	PUBLIC PROPERTY GET monthly_non_accrual
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen monthly_non_accrual, 8, 10, 31
		monthly_non_accrual = trim(monthly_non_accrual)
	END PROPERTY
	
	' NPA arrears
	PUBLIC PROPERTY GET npa_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen npa_arrears, 8, 9, 70
		npa_arrears = trim(npa_arrears)
	END PROPERTY
	
	' PA arrears
	PUBLIC PROPERTY GET pa_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen pa_arrears, 8, 10, 70
		pa_arrears = trim(pa_arrears)
	END PROPERTY
	
	' Total arrears
	PUBLIC PROPERTY GET ttl_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen ttl_arrears, 8, 11, 70
		ttl_arrears = trim(ttl_arrears)
	END PROPERTY		
END CLASS

