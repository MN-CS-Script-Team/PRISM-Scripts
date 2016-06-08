'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - FIND NAME ON CALI.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED

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

' Set up dialog box
'BeginDialog CALI_search_dialog, 0, 0, 216, 115, "CALI Search Criteria"
'  EditBox 55, 25, 75, 15, first_name
'  EditBox 55, 45, 75, 15, last_name
'  DropListBox 50, 80, 90, 15, "Select Unit"+chr(9)+"Child Support 1"+chr(9)+"Child Support 2"+chr(9)+"Child Support 3"+chr(9)+"Child Support 4", Group_dropdown_list
'  EditBox 55, 95, 40, 15, position
'  ButtonGroup ButtonPressed
'    PushButton 140, 35, 65, 15, "Find on your CALI", find_button
'    PushButton 145, 80, 65, 15, "Find on this CALI", find_CALI_button
'    CancelButton 160, 100, 50, 15
'  Text 10, 25, 40, 15, "First Name:"
'  Text 10, 45, 40, 15, "Last Name:"
'  Text 10, 80, 25, 10, "Unit:"
'  Text 10, 95, 35, 10, "Position:"
'  Text 5, 5, 145, 20, "Please enter one or more search criteria for your CALI search:"
'  Text 5, 65, 195, 15, "Enter these optional fields to search another CALI caseload:"
'EndDialog

BeginDialog CALI_search_dialog, 0, 0, 211, 170, "CALI Search Criteria"
  Text 5, 5, 205, 10, "Please enter one or more search criteria for your CALI search:"
  Text 10, 25, 40, 15, "First Name:"
  EditBox 55, 25, 75, 15, first_name
  Text 10, 45, 40, 15, "Last Name:"
  EditBox 55, 45, 75, 15, last_name
  Text 5, 70, 195, 10, "Enter these optional fields to search another CALI caseload:"
  Text 5, 90, 25, 10, "County:"
  EditBox 35, 85, 30, 15, cali_office
  Text 75, 90, 25, 10, "Team:"
  EditBox 105, 85, 25, 15, cali_team
  Text 145, 90, 30, 10, "Position:"
  EditBox 180, 85, 25, 15, cali_position
  ButtonGroup ButtonPressed
    PushButton 140, 45, 65, 15, "Find on your CALI", find_button
    PushButton 140, 110, 65, 15, "Find on this CALI", find_CALI_button
    CancelButton 155, 150, 50, 15
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
'***********************************************************************************************************************************************
'*************************************************************************************
' Custom Function for finding a name in CALI
' Paramaters: name, CALI_unit, CALI_position
' name = name we are searching for
'CALI_office = the office to search
' CALI_team = the "team" or child support unit for the position to be searched
' CALI_position = the position to be searched
'*************************************************************************************
FUNCTION find_name_in_CALI(name, CALI_office, CALI_team, CALI_position)
	EMReadScreen check_for_position_list, 22, 8, 36
		IF check_for_position_list = "Caseload Position List" THEN
			PF3
		END IF
	EMReadScreen check_for_caseload_list, 13, 2, 32
		If check_for_caseload_list = "Caseload List" THEN	
			CALL navigate_to_PRISM_screen("MAIN")
			transmit
		END IF	
	CALL navigate_to_PRISM_screen("CALI")  'Navigate to CALI, remove any case number entered, and display the desired CALI listing
	EMWriteScreen "             ", 20, 58
	EMWriteScreen "  ", 20, 69
	EMWriteScreen CALI_office, 20, 18
	EMWriteScreen "001", 20, 30
	EMWriteScreen CALI_team, 20, 40
	EMWriteScreen CALI_position, 20, 49
	transmit

	name = UCASE(name) 'convert the name that the user entered as search criteria to all caps (or names won't be found!)

	'Set up variables for loop for searching through CALI listing of CP's for the search criteria
	cali_row = 8  'navigates to the first case listed in CALI 
	found = FALSE 
	found_once = FALSE
	DO 
		EMReadScreen end_of_data, 11, cali_row, 32   
		EMReadScreen CP_name, 35, cali_row, 38
		'	msgbox "Search for " & name & " in " & CP_name & "."	

		'If the name we are searching for is in the CALI list of CP's, display a message box to the user to indicate whether 
		'we continue searching for another match.  If the user does not wish to continue searching, the matched case
		'CAST screen is displayed.

		IF INSTR(CP_name, name) > 0 THEN
			EMReadScreen PRISM_number, 10, cali_row, 7
			EMReadScreen case_number, 2, cali_row, 19
			msg = msgbox (name & " is CP on line " & CSTR(cali_row - 7) & " on your CALI list, case number " & PRISM_number & "-" & case_number &".  Continue searching for another match?", 4)
			found_once = TRUE
			IF msg = 7 THEN
				EMWriteScreen "D", cali_row, 3
				transmit		
				found = TRUE 
				stopscript
			END IF 
		END IF 
		IF end_of_data <> "End of Data" THEN
			cali_row = cali_row + 1			
		END IF
		IF cali_row = 19 THEN    'Navigate to a new page 
			cali_row = 8
			PF8
		END IF
	LOOP UNTIL found = TRUE OR end_of_data = "End of Data"

	'Re-set CALI and variables for a second search, this time searching in the CALI list of NCP's.
	EMWriteScreen "             ", 20, 58
	EMWriteScreen "  ", 20, 69
	EMWriteScreen CALI_office, 20, 18
	EMWriteScreen "001", 20, 30
	EMWriteScreen CALI_team, 20, 40
	EMWriteScreen CALI_position, 20, 49
	transmit
	end_of_data = " "
	cali_row = 8  'navigates to the first case listed in CALI 
	found = FALSE
	PF11
	DO 
		EMReadScreen end_of_data, 11, cali_row, 32   
		EMReadScreen NCP_name, 35, cali_row, 33			
		'	msgbox "Search for " & name & " in " & NCP_name & "."

		'If the name we are searching for is in the CALI list of NCP's, display a message box to the user to indicate whether 
		'we continue searching for another match.  If the user does not wish to continue searching, the matched case
		'CAST screen is displayed.
		IF INSTR(NCP_name, name) > 0 THEN
			EMReadScreen PRISM_number, 10, cali_row, 7
			EMReadScreen case_number, 2, cali_row, 19
			msg = msgbox (name & " is NCP on line " & Cstr(cali_row - 7) & " on your CALI list, case number " & PRISM_number & "-" & case_number &".  Continue searching for another match?", 4)
			found_once = TRUE
			IF msg = 7 THEN
				EMWriteScreen "D", cali_row, 3
				transmit
				found = TRUE 
				stopscript
			END IF
		END IF 			
		IF end_of_data <> "End of Data" THEN
			cali_row = cali_row + 1		
		END IF
		IF cali_row = 19 THEN    'Navigate to a new page 
			cali_row = 8
			PF8
		END IF
	LOOP UNTIL found = TRUE OR end_of_data = "End of Data"

' Determine whether any match was found, and display appropriate message.
	IF found_once = TRUE THEN
		msgbox name & " was not found again on your CALI list." 
	ELSE			
		msgbox name & " was not found on your CALI list." 
	END IF 	
END FUNCTION
'**********************************************************************************************
' 
'**********************************************************************************************
EMConnect "" 'Connect to PRISM

DO
	err_msg = ""
	dialog CALI_search_dialog 'Display the dialog 
		CALL check_for_PRISM (false) 'Check to see if PRISM is locked

		IF buttonpressed = 0 THEN stopscript  'If cancel is pressed, end script

		IF buttonpressed = find_button THEN  'The user selected to search on their own CALI
			CALL check_for_PRISM (TRUE)
			CALL navigate_to_PRISM_screen("CALI")  'Navigate to CALI, remove any case number entered
			EMWriteScreen "             ", 20, 58  'and make note of the user's unit and position.
			EMWriteScreen "  ", 20, 69
			transmit
				'Check to see if the user entered a first name, a last name, both or neither
				'If neither, prompt the user to enter search criteria.
				'Otherwise, call the custom function with the appropriate parameters	
			IF LEN(last_name) = 0 and LEN(first_name) = 0 THEN
				DO	
					msgbox "Please enter either a first and/or last name for the search."
					dialog CALI_search_dialog
				LOOP UNTIL LEN(last_name) <> 0 OR LEN(first_name) <> 0
			ELSEIF LEN(last_name) = 0 and LEN(first_name) > 0 THEN
				CALL find_name_in_CALI (first_name, office, unit, position)
			ELSEIF LEN(last_name) > 0 and LEN(first_name) = 0 THEN
				CALL find_name_in_CALI (last_name, office, unit, position)
			ELSEIF LEN(last_name) > 0 and LEN(first_name) > 0 THEN
				CALL find_name_in_CALI (last_name & ", " & first_name, office, unit, position)
			END IF 
			script_end_procedure("")
		ELSEIF ButtonPressed = find_CALI_button THEN  'The user selected to search a specific CALI listing
			CALL check_for_PRISM (false)

			IF first_name = "" AND last_name = "" THEN err_msg = err_msg & vbCr & "* Please enter a first and/or last name."
			IF len(CALI_position) <> 2 THEN err_msg = err_msg & vbCr & "* Please enter a valid, 2-digit position number."
			IF IsNumeric(CALI_position) = FALSE THEN err_msg = err_msg & vbCr & "* Please enter a valid, 2-digit position number."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		END IF
LOOP UNTIL err_msg = ""			
		
CALL navigate_to_PRISM_screen("CALI")  'Navigate to CALI, remove any case number entered
EMWriteScreen "             ", 20, 58  'and display the desired CALI listing
EMWriteScreen "  ", 20, 69
transmit

'Check to see if the user entered a first name, a last name, or both.
     'Call the custom function with the appropriate parameters.	
IF LEN(last_name) = 0 and LEN(first_name) > 0 THEN
	CALL find_name_in_CALI (first_name, CALI_office, CALI_team, CALI_position)
ELSEIF LEN(last_name) > 0 and LEN(first_name) = 0 THEN
	CALL find_name_in_CALI (last_name, CALI_office, CALI_team, CALI_position)
ELSEIF LEN(last_name) > 0 and LEN(first_name) > 0 THEN
	CALL find_name_in_CALI (last_name & ", " & first_name, CALI_office, CALI_team, CALI_position)
END IF 
	
script_end_procedure("")
