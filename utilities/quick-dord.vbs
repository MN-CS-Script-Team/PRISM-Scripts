'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "quick-dord.vbs"
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

script_end_procedure("This script will be released sometime, no ETA. - VKC")

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/23/2016", "This script has been moved to the utilities category. End users should not notice any changes.", "Veronica Cary, DHS")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'This is a special function which writes to the text file used by this script
Function write_line_to_text_file(new_line_for_writing, file_location)
	'Now it determines the favorite CAAD notes
	With (CreateObject("Scripting.FileSystemObject"))							'Creating an FSO
		If .FileExists(user_myDocs_folder & "favoriteCAADnotes.txt") Then		'If the file exists...
			Set TextFileObj = CreateObject("Scripting.FileSystemObject")		'Create another FSO
			Set text_command = TextFileObj.OpenTextFile(file_location)			'Open the text file
			text_raw = text_command.ReadAll										'Read the text file
			text_command.Close													'Closes the file
		END IF
	END WITH

	'Now it updates the text file
	With (CreateObject("Scripting.FileSystemObject"))							'Creating an FSO
		If .FileExists(file_location) Then create_new_file = true				'Setting this variable now so the script can apply the logic later
		Set TextFileObj = CreateObject("Scripting.FileSystemObject")			'Create another FSO
		Set text_command = TextFileObj.OpenTextFile(file_location, 2, True)		'Open the text file for writing
		If create_new_file = true Then											'If the file existed, it should write as a new line (this avoids a vbNewLine being entered as the first item in the file)
			text_command.Write text_raw & vbNewLine & new_line_for_writing 		'Writes the new line in an existing file
		Else																	'If the file doesn't exist, it should simply add the new details without a new line
			text_command.Write new_line_for_writing 							'Writes the new in a new file
		End if																	'End of if...then statement
		text_command.Close														'Closes the file
	END WITH
End function

'This is a special function which clears a line from the text file used by this script
Function clear_line_from_text_file(line_for_clearing, file_location)
	'Now it determines the favorite CAAD notes
	With (CreateObject("Scripting.FileSystemObject"))							'Creating an FSO
		If .FileExists(user_myDocs_folder & "favoriteCAADnotes.txt") Then		'If the file exists...
			Set TextFileObj = CreateObject("Scripting.FileSystemObject")		'Create another FSO
			Set text_command = TextFileObj.OpenTextFile(file_location)			'Open the text file
			text_raw = text_command.ReadAll										'Read the text file
			text_command.Close													'Closes the file
		END IF
	END WITH

	text_filtered = replace(text_raw, line_for_clearing & vbNewLine, "")

	'Now it updates the text file
	With (CreateObject("Scripting.FileSystemObject"))							'Creating an FSO
		Set TextFileObj = CreateObject("Scripting.FileSystemObject")			'Create another FSO
		Set text_command = TextFileObj.OpenTextFile(file_location, 2)			'Open the text file for writing
		text_command.Write text_filtered			 							'Writes the new in a new file
		text_command.Close														'Closes the file
	END WITH
End function

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

'Sets location of files as the user_myDocs_folder from above
file_location = user_myDocs_folder & "favoriteCAADnotes.txt"


'This is a large do...loop which will constantly re-read the text file in case changes were made. If the script is just being used (and not modified), it will only iterate once through this.
Do

	'Before loading dialogs, needs to scan the My Documents folder for a file called favoriteCAADnotes.txt. If this file is found, details about CAAD notes will be pre-loaded. Otherwise, it won't be.
	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now it determines the favorite CAAD notes
	With (CreateObject("Scripting.FileSystemObject"))																				'Creating an FSO
		If .FileExists(user_myDocs_folder & "favoriteCAADnotes.txt") Then															'If the favoriteCAADnotes.txt file exists...
			Set get_favorite_CAAD_notes = CreateObject("Scripting.FileSystemObject")												'Create another FSO
			Set favorite_CAAD_notes_command = get_favorite_CAAD_notes.OpenTextFile(user_myDocs_folder & "favoriteCAADnotes.txt")	'Open the text file
			favorite_CAAD_notes_raw = favorite_CAAD_notes_command.ReadAll															'Read the text file <<<<<CAN THIS READ ONE LINE
			IF favorite_CAAD_notes_raw <> "" THEN favorite_CAAD_notes_array = split(favorite_CAAD_notes_raw, vbNewLine)				'Split by new lines
			favorite_CAAD_notes_command.Close																						'Closes the file
		Else																														'...if the file doesn't exist...
			MsgBox "Welcome to the Quick CAAD script! This appears to be your first time running this." & vbNewLine & vbNewLine & _
				"You can use this script to store several CAAD codes you may use frequently. To start, navigate to the search screen and search for a CAAD code to add."
				favorite_CAAD_notes_array = array()
		End if
	END WITH

	'THIS IS A DYNAMIC DIALOG! DON'T OPEN ME IN DIALOG EDITOR!!!!
	BeginDialog quick_CAAD_dialog, 0, 0, 300, 200, "Quick CAAD dialog"
	  ButtonGroup ButtonPressed
	    OkButton 190, 180, 50, 15
	    CancelButton 245, 180, 50, 15
		dialog_row = 5																										'Starting here so that the contents display in a pretty manner
		'Iterate through each item in the array determined above, then display them
		For i = 0 to ubound(favorite_CAAD_notes_array)																		'i is a counter in this case
			If favorite_CAAD_notes_array(i) = "" then exit for 																'If it's blank, we should stop because we're likely at the end of the file
			number_to_pass_to_the_button = 1000 + i																			'Add 1000 to the counter to get a ButtonPressed value we can use
			PushButton 5, dialog_row, 30, 10, left(favorite_CAAD_notes_array(i), 5), number_to_pass_to_the_button			'Show the 5-most characters of the CAAD code (each one is five characters) as a button
			Text 40, dialog_row, 255, 10, right(favorite_CAAD_notes_array(i), len(favorite_CAAD_notes_array(i)) - 7)		'Show the rest as a description
			dialog_row = dialog_row + 15																					'Go up 15 pixels
		Next
		PushButton 5, 185, 80, 10, "search CAAD codes...", search_CAAD_codes_button											'Provides a search feature
		PushButton 90, 185, 80, 10, "delete saved codes...", delete_saved_codes_button										'Provides a delete feature
	EndDialog

	'THIS IS A DYNAMIC DIALOG! DON'T OPEN ME IN DIALOG EDITOR!!!!
	BeginDialog quick_CAAD_delete_codes_dialog, 0, 0, 300, 200, "Quick CAAD: delete codes dialog"
	  ButtonGroup ButtonPressed
		OkButton 245, 180, 50, 15
		dialog_row = 5																										'Starting here so that the contents display in a pretty manner
		'Iterate through each item in the array determined above, then display them
		For i = 0 to ubound(favorite_CAAD_notes_array)																		'i is a counter in this case
			If favorite_CAAD_notes_array(i) = "" then exit for 																'If it's blank, we should stop because we're likely at the end of the file
			number_to_pass_to_the_button = 1000 + i																			'Add 1000 to the counter to get a ButtonPressed value we can use
			PushButton 5, dialog_row, 30, 10, "delete", number_to_pass_to_the_button										'Show the delete button
			Text 40, dialog_row, 255, 10, favorite_CAAD_notes_array(i)														'Show the description
			dialog_row = dialog_row + 15																					'Go up 15 pixels
		Next
	EndDialog

	BeginDialog quick_CAAD_search_dialog, 0, 0, 360, 120, "Quick CAAD search dialog"
	EditBox 115, 15, 55, 15, CAAD_code_to_search
	EditBox 115, 50, 105, 15, CAAD_description_to_search
	  ButtonGroup ButtonPressed
	    OkButton 245, 100, 50, 15
	    CancelButton 300, 100, 50, 15
	  GroupBox 5, 5, 170, 30, "Option A"
	  Text 10, 20, 100, 10, "Enter a CAAD code to search:"
	  GroupBox 5, 40, 220, 30, "Option B"
	  Text 10, 55, 100, 10, "Enter a description to search:"
	  GroupBox 230, 5, 120, 90, "Instructions"
	  Text 235, 15, 110, 75, "Enter a five-digit CAAD code (ex: M9901) into Option A, or (alternatively) enter a description to search for in Option B. If a description is searched, the script will return any matches one by one, and allow you to indicate whether-or-not it is the correct CAAD code."
	EndDialog

	'Connects to PRISM
	EMConnect ""

	'Displays the dialog, cancels if cancel is pressed
	Dialog quick_CAAD_dialog
	If ButtonPressed = cancel then StopScript

	'Iterates through the array of favorite CAAD notes, and compares with the ButtonPressed value... If the ButtonPressed = i + 1000 (see above), it will have found the right CAAD code!
	For i = 0 to ubound(favorite_CAAD_notes_array)
		If ButtonPressed = i + 1000 then CAAD_code_to_write = left(favorite_CAAD_notes_array(i), 5)	'If the ButtonPressed = i + 1000, then we found the right CAAD note, and can get the CAAD_code_to_write from the left-most 5 characters
	Next

	'If the user wants to search, we need to guide them toward that. Here's how...
	If ButtonPressed = search_CAAD_codes_button then

		'This do...loop will run until ButtonPressed is cancel, or until the user writes info in either CAAD_code_to_search or CAAD_description_to_search.
		Do
			Dialog quick_CAAD_search_dialog																												'Show the dialog
			If ButtonPressed = cancel then exit do																										'Exit do (so we can get out to the main dialog)
			If CAAD_code_to_search = "" and CAAD_description_to_search = "" then MsgBox "You must enter either a CAAD code or description to search."	'Error message

			'Checks for PRISM (password out) before we continue
			call check_for_PRISM(true)

			'If the CAAD_code_to_search is entered, this will process that first. Otherwise it will go through the description to search.
			If CAAD_code_to_search <> "" then
				navigate_to_PRISM_screen("CAAD")																										'Gets to CAAD
				PF5																																		'Gets to the "add" menu
				EMSetCursor 4, 54																														'Where the code is entered, we need to set the cursor there to read the help details
				PF1																																		'Loads help
				EMWriteScreen CAAD_code_to_search, 20, 28																								'Write the search string at 20, 28 on the screen
				transmit																																'Transmits
				EMReadScreen CAAD_code_check, 5, 13, 18																									'Checks the first response
				If CAAD_code_check = CAAD_code_to_search then																							'If the first response is correct...
					add_to_quick_CAAD = MsgBox("The code was found! Would you like to add this to the Quick CAAD button?", vbYesNo + vbQuestion)		'...ask the user if they want to update...
					If add_to_quick_CAAD = vbNo then exit do																							'...If they don't, take them back to the main dialog...
					If add_to_quick_CAAD = vbYes then																									'...If they do, time to add the file!
						EMReadScreen CAAD_description_to_add, 54, 13, 24																				'Reads the description for good measure
						call write_line_to_text_file(CAAD_code_to_search & ", " & CAAD_description_to_add, file_location)								'Uses custom function to write to the file
						PF3																																'Gets out of the screen!
					End if
				Else
					MsgBox "Your CAAD code was not found."		'Yells at you
					PF3											'Gets out of the screen!
					CAAD_code_to_search = ""					'Blanks this out so we don't leave the search dialog
				End if
			ElseIf CAAD_description_to_search <> "" then
				navigate_to_PRISM_screen("CAAD")																																	'Gets to CAAD
				PF5																																									'Gets to the "add" menu
				EMSetCursor 4, 54																																					'Where the code is entered, we need to set the cursor there to read the help details
				PF1																																									'Loads help
				For PRISM_row = 13 to 19																																			'Available rows in PRISM for reading
					EMReadScreen CAAD_code_check, 5, PRISM_row, 18																													'Checks the code
					EMReadScreen CAAD_description_check, 55, PRISM_row, 24																											'Checks the description
					If instr(CAAD_description_check, ucase(CAAD_description_to_search)) > 0 then																					'If the search string is found in the description then it generates a notice
						match_notice = MsgBox("A match has been found!" & vbNewLine & vbNewLine & _
						"Is CAAD code " & CAAD_code_check & ": " & trim(CAAD_description_check) & " what you're looking for?" & vbNewLine & vbNewLine & _
						"Press ''Yes'' to add this to your Quick CAAD list, ''No'' to keep searching, or ''Cancel'' to return to the primary menu.", vbYesNoCancel + vbQuestion)	'Yes adds it, No keeps searching, cancel returns to the menu
						If match_notice = vbCancel then Exit For																													'Return to the menu
						If match_notice = vbYes then																																'If yes...
							call write_line_to_text_file(CAAD_code_check & ", " & trim(CAAD_description_check), file_location)														'Uses custom function to write to the file
							Exit for																																				'Return to the menu
						End if
					End if
					EMReadScreen end_of_data_check, 19, 14, 36																	'Are we out of data to scan?
					If end_of_data_check = "*** End of Data ***" then															'If so, we should alert the worker then exit this section and return to the main menu
						MsgBox "No match found!"
						Exit for
					End if
					If PRISM_row = 19 then																						'If we're at the end, it should reset to row 13 so it keeps looping
						PRISM_row = 13
						PF8																										'Gets to next screen to keep searching
					End if
				Next
			End if
		Loop until ButtonPressed = cancel or (CAAD_code_to_search <> "" or CAAD_description_to_search <> "")

		ButtonPressed = 0	'Needs to blank out so it goes back to the main dialog

	ElseIf ButtonPressed = delete_saved_codes_button then
		Do
			Dialog quick_CAAD_delete_codes_dialog
			If ButtonPressed = OK then exit do

			'Iterates through the array of favorite CAAD notes, and compares with the ButtonPressed value... If the ButtonPressed = i + 1000 (see above), it will have found the right CAAD code, and will clear that string from the text file
			For i = 0 to ubound(favorite_CAAD_notes_array)
				If ButtonPressed = i + 1000 then
					quick_CAAD_warning_box = MsgBox("Are you sure you want to clear this one?", vbOKCancel + vbQuestion)
					If quick_CAAD_warning_box = vbOK then call clear_line_from_text_file(favorite_CAAD_notes_array(i), file_location)
				End if
			Next

		Loop until quick_CAAD_warning_box <> vbCancel

		ButtonPressed = 0	'Needs to blank out so it goes back to the main dialog

	End if

Loop until ButtonPressed <> cancel

'Checks for PRISM (password out) before we continue
call check_for_PRISM(true)

'Now it goes to the selected note
navigate_to_PRISM_screen("CAAD")			'Navigates to CAAD
PF5											'Adds a new note
EMWriteScreen CAAD_code_to_write, 4, 54		'Writes the code
