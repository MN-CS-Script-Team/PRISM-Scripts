'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "favorites-list.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/04/2017", "The favorites script is now ready for use!", "Veronica Cary, DHS")
call changelog_update("02/22/2017", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'INSTRUCTIONS BLOCK ========================================================================================================
'~~~~~ This script either creates a new favorites list for your use, or allows you to view and run your favorites scripts, as well as new scripts.
'~~~~~ On the first start-up, you'll be prompted to select your favorite scripts from a list. The script will then exit in this case, and write a text file to your My-Documents folder containing your favorite scripts.
'~~~~~ After this, for general use, simply start the script, then select the script you want to run. That's it!
'+++++ https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Image%20Library/ACTIONS%20-%20ABAWD%20FIATER/initial%20dialog.png
'***** This is a sample image and you need to sign in to SIR to see it. It isn't related to this script at all, but is just to test the process.
'~~~~~ To edit your list, simply select the "update favorites" button at the bottom of the window.
'===== Keywords: favorites, customization
'END INSTRUCTIONS BLOCK ====================================================================================================

'DEFINING SOME VARIABLES ===================================================================================================
button_height = 12
button_width = 145
dialog_margin = 5
groupbox_margin = 5
'END VARIABLES =============================================================================================================

'This function simply displays a list of hotkeys, and the user can insert screens-to-navigate-to within
function edit_hotkeys
	'Instructional MsgBox
	MsgBox  "This section will add PRISM screens to hotkey combinations!" & vbNewLine & vbNewLine & _
			"To use it, simply insert the four-character PRISM screen you'd like to navigate to when pressing the specific key combination." & vbNewLine & vbNewLine & _
			"So, for example, to navigate to CAAD every time Ctrl-F1 is pressed, simply type ""CAAD"" in the editbox." & vbNewLine & vbNewLine & _
			"When you are finished, the script will add a hotkeys file to your My Documents folder, which will store your choices."

	'A dialog
	BeginDialog hotkey_selection_dialog, 0, 0, 116, 285, "Hotkey Selection Dialog"
	  Text 15, 10, 30, 10, "Hotkey:"
	  Text 55, 5, 55, 20, "PRISM screen to navigate to:"
	  Text 15, 30, 25, 10, "Ctrl-F1:"
	  Text 15, 50, 25, 10, "Ctrl-F2:"
	  Text 15, 70, 25, 10, "Ctrl-F3:"
	  Text 15, 90, 25, 10, "Ctrl-F4:"
	  Text 15, 110, 25, 10, "Ctrl-F5:"
	  Text 15, 130, 25, 10, "Ctrl-F6:"
	  Text 15, 150, 25, 10, "Ctrl-F7:"
	  Text 15, 170, 25, 10, "Ctrl-F8:"
	  Text 15, 190, 25, 10, "Ctrl-F9:"
	  Text 10, 210, 30, 10, "Ctrl-F10:"
	  Text 10, 230, 30, 10, "Ctrl-F11:"
	  Text 10, 250, 30, 10, "Ctrl-F12:"
	  EditBox 55, 25, 55, 15, ctrl_f1_hotkey_choice
	  EditBox 55, 45, 55, 15, ctrl_f2_hotkey_choice
	  EditBox 55, 65, 55, 15, ctrl_f3_hotkey_choice
	  EditBox 55, 85, 55, 15, ctrl_f4_hotkey_choice
	  EditBox 55, 105, 55, 15, ctrl_f5_hotkey_choice
	  EditBox 55, 125, 55, 15, ctrl_f6_hotkey_choice
	  EditBox 55, 145, 55, 15, ctrl_f7_hotkey_choice
	  EditBox 55, 165, 55, 15, ctrl_f8_hotkey_choice
	  EditBox 55, 185, 55, 15, ctrl_f9_hotkey_choice
	  EditBox 55, 205, 55, 15, ctrl_f10_hotkey_choice
	  EditBox 55, 225, 55, 15, ctrl_f11_hotkey_choice
	  EditBox 55, 245, 55, 15, ctrl_f12_hotkey_choice
	  ButtonGroup ButtonPressed
	    OkButton 5, 265, 50, 15
	    CancelButton 60, 265, 50, 15
	EndDialog

	'Show the dialog
	Dialog hotkey_selection_dialog
	If ButtonPressed = cancel then StopScript

	'>>> If the user has already selected their hotkeys, the script will open that file and
	'>>> and read it, storing the contents in the variable name ''favorites_text_file_array''
	SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
	With oTxtFile
		If .FileExists(hotkeys_text_file_location) Then
			Set hotkeys = CreateObject("Scripting.FileSystemObject")
			Set hotkeys_command = hotkeys.OpenTextFile(hotkeys_text_file_location)
			hotkeys_array = hotkeys_command.ReadAll
			hotkeys_command.Close
		Else
			MsgBox "file not found"
			Set hotkeys = CreateObject("Scripting.FileSystemObject")
			Set hotkeys_command = hotkeys.CreateTextFile(hotkeys_text_file_location)
			hotkeys_command.Write("Hey")
			hotkeys_command.Close
		End if
	END WITH

	'Somehow program the redirects to look at that file and do the magic

end function

'====================================================================================
'====================================================================================
'This VERY VERY long function contains all of the logic behind editing the favorites.
'====================================================================================
'====================================================================================
function edit_favorites

	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>> SECTION 1 <<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that happen before the user sees anything. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


	'Looks up the script details online (or locally if you're a scriptwriter)

	If run_locally <> true then

		'Creating the object to the URL a la text file
		SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")

		'Building an array of all scripts
		'Opening the URL for the given main menu
		get_all_scripts.open "GET", all_scripts_repo, FALSE
		get_all_scripts.send
		IF get_all_scripts.Status = 200 THEN
			Set filescriptobject = CreateObject("Scripting.FileSystemObject")
			Execute get_all_scripts.responseText
		ELSE
			'If the script cannot open the URL provided...
			MsgBox 	"Something went wrong with the URL: " & all_scripts_repo
			stopscript
		END IF
	ELSE
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(all_scripts_repo)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF

	'Warning/instruction box
	MsgBox  "This section will display a dialog with various scripts on it. Any script you check will be added to your favorites menu. Scripts you un-check will be removed. Once you are done making your selection hit ""OK"" and your menu will be updated. " & vbNewLine & vbNewLine &_
			"Note: you will be unable to edit the list of new scripts."

	'An array containing details about the list of scripts, including how they are displayed and stored in the favorites tag
	'0 => The script name
	'1 => The checked/unchecked status (based on the dialog list)
	'2 => The script category, and a "/" so that it's presented in a URL
	'3 => The proper script file name
	'4 => The hotkey the user has associated with the script

	REDIM scripts_edit_favs_array(ubound(cs_scripts_array), 4)

	'determining the number of each kind of script...by category
	number_of_scripts = 0
	actions_scripts = 0
	bulk_scripts = 0
	calc_scripts = 0
	notes_scripts = 0
	utilities_scripts = 0
	FOR i = 0 TO ubound(cs_scripts_array)
		number_of_scripts = i
		IF cs_scripts_array(i).category = "actions" THEN
			actions_scripts = actions_scripts + 1
		ELSEIF cs_scripts_array(i).category = "bulk" THEN
			bulk_scripts = bulk_scripts + 1
		ELSEIF cs_scripts_array(i).category = "calculators" THEN
			calc_scripts = calc_scripts + 1
		ELSEIF cs_scripts_array(i).category = "notes" THEN
			notes_scripts = notes_scripts + 1
		ELSEIF cs_scripts_array(i).category = "utilities" THEN
	        utilities_scripts = utilities_scripts + 1
	    End if
	NEXT


	'>>> If the user has already selected their favorites, the script will open that file and
	'>>> and read it, storing the contents in the variable name ''favorites_text_file_array''
	SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
	With oTxtFile
		If .FileExists(favorites_text_file_location) Then
			Set fav_scripts = CreateObject("Scripting.FileSystemObject")
			Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
			fav_scripts_array = fav_scripts_command.ReadAll
			IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
			fav_scripts_command.Close
		END IF
	END WITH

	'>>> Determining the width of the dialog from the number of scripts that are available...
	'the dialog starts with a width of 400
	dia_width = 400

	'VKC - removed old functionality to determine dynamically the width. This will need to be redetermined based on the number of scripts, but I am holding off on this until I know all of the content I'll jam in here. -11/29/2016

	'>>> Building the dialog
	BeginDialog build_new_favorites_dialog, 0, 0, dia_width, 440, "Select your favorites"
		ButtonGroup ButtonPressed
			OkButton 5, 5, 50, 15
			CancelButton 55, 5, 50, 15
			PushButton 165, 5, 70, 15, "Reset Favorites", reset_favorites_button
		'>>> Creating the display of all scripts for selection (in checkbox form)
		script_position = 0		' <<< This value is tied to the number_of_scripts variable


		col = 10
		row = 30
		Text col, row, 175, 10, "---------- ACTIONS SCRIPTS ----------"
		row = row + 10

		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "actions" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- BULK SCRIPTS ----------"
		row = row + 10

		'BULK script laying out
		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "bulk" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- CALCULATOR SCRIPTS ----------"
		row = row + 10

		'CALCULATOR script laying out
		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "calculators" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- NOTES SCRIPTS ----------"
		row = row + 10

		'NOTES script laying out
		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "notes" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- UTILITIES SCRIPTS ----------"
		row = row + 10

		'UTILITIES script laying out
	    FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "utilities" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT
	EndDialog

	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>> SECTION 2 <<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that the user sees and makes do. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<

	DO
		DO
			'>>> Running the dialog
			Dialog build_new_favorites_dialog
			'>>> Cancel confirmation
			IF ButtonPressed = 0 THEN
				confirm_cancel = MsgBox("Are you sure you want to cancel? Press YES to cancel the script. Press NO to return to the script.", vbYesNo)
				IF confirm_cancel = vbYes THEN script_end_procedure("~PT: Script cancelled.")
			END IF
			'>>> If the user selects to reset their favorites selections, the script
			'>>> will go through the multi-dimensional array and reset all the values
			'>>> for position 1, thereby clearing the favorites from the display.
			IF ButtonPressed = reset_favorites_button THEN
				FOR i = 0 to number_of_scripts
					scripts_edit_favs_array(i, 1) = unchecked
				NEXT
			END IF
		'>>> The exit condition for the first do/loop is the user pressing 'OK'
		LOOP UNTIL ButtonPressed <> 0 AND ButtonPressed <> reset_favorites_button
		'>>> Validating that the user does not select more than a prescribed number of scripts.
		'>>> Exceeding the limit will cause an exception access violation for the Favorites script when it runs.
		'>>> Currently, that value is 30. That is lower than previous because of the larger number of new scripts. (-Robert, 04/20/2016)
		double_check_array = ""
		FOR i = 0 to number_of_scripts
			IF scripts_edit_favs_array(i, 1) = checked THEN double_check_array = double_check_array & scripts_edit_favs_array(i, 0) & "~"
		NEXT
		double_check_array = split(double_check_array, "~")
		IF ubound(double_check_array) > 29 THEN MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 30."
		'>>> Exit condition is the user having fewer than 30 scripts in their favorites menu.
	LOOP UNTIL ubound(double_check_array) <= 29

	'>>> Getting ready to write the user's selection to a text file and save it on a prescribed location on the network.
	'>>> Building the content of the text file.
	FOR i = 0 to number_of_scripts - 1
		IF scripts_edit_favs_array(i, 1) = checked THEN favorite_scripts = favorite_scripts & scripts_edit_favs_array(i, 2) & scripts_edit_favs_array(i, 3) & vbNewLine
	NEXT

	'>>> After the user selects their favorite scripts, we are going to write (or overwrite) the list of scripts
	'>>> stored at H:\my favorite scripts.txt.
	IF favorite_scripts <> "" THEN
		SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
		SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(favorites_text_file_location, 2)
		updated_fav_scripts_command.Write(favorite_scripts)
		updated_fav_scripts_command.Close
		script_end_procedure("Success!! Your Favorites Menu has been updated. Please click your favorites list button to re-load them.")
	ELSE
		'>>> OR...if the user has selected no scripts for their favorite, the file will be deleted to
		'>>> prevent the Favorites Menu from erroring out.
		'>>> Experience with worker_signature automation tells us that if the text file is blank, the favorites menu doth not work.
		oTxtFile.DeleteFile(favorites_text_file_location)
		script_end_procedure("You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites.")
	END IF

end function

'>>> Determining the location of the user's favorites list.

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

favorites_text_file_location = user_myDocs_folder & "\scripts-cs-favorites.txt"
hotkeys_text_file_location = user_myDocs_folder & "\scripts-cs-hotkeys.txt"

'switching up the script_repository because the all scripts file is not in the Script Files folder
all_scripts_repo = script_repository & "/~complete-list-of-scripts.vbs"

'========================================================================================================
'========================================================================================================
'========================================================================================================
'========================================================================================================
'================================================================================== NOW THE ACTUAL SCRIPT
'========================================================================================================
'========================================================================================================
'========================================================================================================
'========================================================================================================

'>>> favorited_scripts_array will be built from the contents of the user's text file
favorited_scripts_array = ""

'Does this differently if you're a run_locally user vs not
If run_locally <> true then
	'>>> Creating the object needed to connect to the interwebs.
	SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")
	get_all_scripts.open "GET", all_scripts_repo, FALSE
	get_all_scripts.send
	IF get_all_scripts.Status = 200 THEN
		Set filescriptobject = CreateObject("Scripting.FileSystemObject")
		Execute get_all_scripts.responseText
	ELSE
		'>>> Displaying the error message when the script fails to connect to a specific main menu.
		'>>> the replace & right bits are there to display the main menu in a way that is clear to the user.
		'>>> We are going to display the right length minus 99 because there are 99 characters between the start of the https and the last / before the main menu name.
		'>>> That length needs to be updated when we go state-wide.
		MsgBox("Something went wrong grabbing trying to locate All Scripts File. Please contact scripts administrator.")
		stopscript
	END IF
ELSE
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(all_scripts_repo)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF

'>>> Building the array of new scripts
num_of_new_scripts = 0

'>>> Looking through the scripts array and determining all of the new ones.
FOR i = 0 TO Ubound(cs_scripts_array)
	IF DateDiff("D", cs_scripts_array(i).release_date, date) < 60 THEN 				
		num_of_new_scripts = num_of_new_scripts + 1		
		ReDim Preserve new_cs_scripts_array(num_of_new_scripts)
		SET new_cs_scripts_array(num_of_new_scripts) = NEW cs_script
		new_cs_scripts_array(num_of_new_scripts).script_name		= cs_scripts_array(i).script_name		
		new_cs_scripts_array(num_of_new_scripts).category			= cs_scripts_array(i).category			
		new_cs_scripts_array(num_of_new_scripts).description		= cs_scripts_array(i).description		
		new_cs_scripts_array(num_of_new_scripts).release_date		= cs_scripts_array(i).release_date		
		new_cs_scripts_array(num_of_new_scripts).scriptwriter		= cs_scripts_array(i).scriptwriter				
	end if
NEXT

'>>> This handles what happens if there are no new scripts (it'll happen)
if num_of_new_scripts = 0 then
	num_of_new_scripts = 1
	ReDim Preserve new_cs_scripts_array(num_of_new_scripts)
	SET new_cs_scripts_array(num_of_new_scripts) = NEW cs_script
	new_cs_scripts_array(num_of_new_scripts).script_name		= "no new scripts found."
	new_cs_scripts_array(num_of_new_scripts).category			= "none"
	new_cs_scripts_array(num_of_new_scripts).description		= "none"
	new_cs_scripts_array(num_of_new_scripts).release_date		= "none"
	new_cs_scripts_array(num_of_new_scripts).scriptwriter		= "none"
end if

'>>> Custom function that builds the Favorites Main Menu dialog.
'>>> the array of the user's scripts
FUNCTION favorite_menu(favorites_text_file_array, script_location)
	'>>> Splitting the array of all scripts.
	favorites_text_file_array = trim(favorites_text_file_array)
	favorites_text_file_array = split(favorites_text_file_array, vbNewLine)

	num_of_user_scripts = ubound(favorites_text_file_array)
	

	num_of_scripts = num_of_user_scripts + num_of_new_scripts

	ReDim favorited_scripts_array(num_of_scripts, 6)
	'position 0 = script name
	'position 1 = script directory
	'position 2 = button
	'position 3 = category
	'position 4 = script name without category
	'position 5 = state-supported true/false
	'position 6 = friendly name

	scripts_pos = 0
	FOR EACH script_path IN favorites_text_file_array
		IF script_path <> "" THEN
			favorited_scripts_array(scripts_pos, 0) = script_path
			'>>> Creating the correct URL for the github call
			IF left(script_path, 5) = "notes" THEN
				favorited_scripts_array(scripts_pos, 1) = script_path
				favorited_scripts_array(scripts_pos, 3) = "NOTES"
				favorited_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 6)
				favorited_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_path, 7) = "actions" THEN
				favorited_scripts_array(scripts_pos, 1) = script_path
				favorited_scripts_array(scripts_pos, 3) = "ACTIONS"
				favorited_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 8)
				favorited_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_path, 4) = "bulk" THEN
				favorited_scripts_array(scripts_pos, 1) = script_path
				favorited_scripts_array(scripts_pos, 3) = "BULK"
				favorited_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 5)
				favorited_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_path, 11) = "calculators" THEN
				favorited_scripts_array(scripts_pos, 1) = script_path
				favorited_scripts_array(scripts_pos, 3) = "CALCULATORS"
				favorited_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 12)
				favorited_scripts_array(scripts_pos, 5) = true
            ELSEIF left(script_path, 9) = "utilities" THEN
    			favorited_scripts_array(scripts_pos, 1) = script_path
    			favorited_scripts_array(scripts_pos, 3) = "UTILITIES"
    			favorited_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 10)
    			favorited_scripts_array(scripts_pos, 5) = true
			END IF
			
			'This part reads the complete list of scripts, and then stores the "friendly name" as an item in the array (makes the dialog prettier down the road)
			for each cs_script_data_from_complete_list in cs_scripts_array
				if favorited_scripts_array(scripts_pos, 4) = cs_script_data_from_complete_list.file_name then favorited_scripts_array(scripts_pos, 6) = cs_script_data_from_complete_list.script_name
			next
			
			scripts_pos = scripts_pos + 1
		END IF
	NEXT

	'>>> Determining the height parameters to enable the group boxes.
	actions_count = 0
	bulk_count = 0
	calc_count = 0
	notes_count = 0
	utilities_count = 0
	FOR i = 0 TO (ubound(favorites_text_file_array) - 1)
		IF favorited_scripts_array(i, 3) = "ACTIONS" THEN
			actions_count = actions_count + 1
		ELSEIF favorited_scripts_array(i, 3) = "BULK" THEN
			bulk_count = bulk_count + 1
		ELSEIF favorited_scripts_array(i, 3) = "CALCULATORS" THEN
			calc_count = calc_count + 1
		ELSEIF favorited_scripts_array(i, 3) = "NOTES" THEN
			notes_count = notes_count + 1
        ELSEIF favorited_scripts_array(i, 3) = "UTILITIES" THEN
    		utilities_count = utilities_count + 1
		END IF
	NEXT

	'>>> Determining the height of the dialog.
	'>>> Each groupbox will require a minimum of 25 pixels. That is the height of the groupbox with 1 script PushButton
	'>>> The groupboxes need to grow 10 for each script pushbutton, so the dialog also needs to grow 10 for each script push button. However,
	'>>> 	the size of each groupbox will always be 15 plus (10 times the number of that kind of script)...
	dlg_height = 0
	IF actions_count <> 0 THEN dlg_height = 15 + (button_height * actions_count)
	IF bulk_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * bulk_count))
	IF calc_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * bulk_count))
	IF notes_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * notes_count))
    IF utilities_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * utilities_count))
	dlg_height = dlg_height + 5
	'>>> The dialog needs to be at least 185 pixels tall. If it is not...because the user has not selected a sufficient number of scripts...then
	'>>> the dialog needs to grow to 185.

	'>>> Adjusting the height if the user has fewer scripts selected (in the left column) than what is "new" (in the right column).
	right_col_dlg_height = 60 + (button_height * (Ubound(new_cs_scripts_array) + 1))
	IF right_col_dlg_height > dlg_height THEN dlg_height = right_col_dlg_height

	
	
	'>>> Defining some variables for use in the display
	groupbox_y_pos = dialog_margin
	left_col_y_pos = groupbox_y_pos + (groupbox_margin * 2)
	left_col_x_pos = dialog_margin + groupbox_margin 
	
	

	'>>> A nice decoration for the user. If they have used Update Worker Signature, then their signature is built into the dialog display.
	IF worker_signature <> "" THEN
		dlg_name = worker_signature & "'s Favorite Scripts"
	ELSE
		dlg_name = "My Favorite Scripts"
	END IF

	'>>> The dialog
	BeginDialog favorites_dialog, 0, 0, 411, dlg_height, dlg_name & " "
  	  ButtonGroup ButtonPressed
		
		'>>> User's favorites
		'>>> This iterates through an array to display the scripts from the favorites text file, in buttons which can be pressed and will run the script.
		
		'Defining these variables before the start of the loop
		number_of_scripts_in_this_category = 1			
		button_placeholder = 100
		
		'The actual array (this goes through the text file and creates scripts and buttons)
		FOR i = 0 TO (ubound(favorites_text_file_array) - 1)		
						
			'Defines the current category for comparison purposes, and to write out the labels.
			current_script_category_from_list = favorited_scripts_array(i, 3)
				
			'Determines the next script category, but only if it's not at the end of the array (because then we're out and the ubound would error out)
			if i + 1 < ubound(favorited_scripts_array) - 1 then
				next_script_category_from_list = favorited_scripts_array(i + 1, 3)
			end if
			
			'Adding the button
			PushButton left_col_x_pos + groupbox_margin, left_col_y_pos, button_width, button_height, favorited_scripts_array(i, 6), button_placeholder
			button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100
			
			'If the current category differs from the next, it's time to make the groupbox
			if current_script_category_from_list <> next_script_category_from_list then
				left_col_y_pos = groupbox_y_pos + ((groupbox_margin * 3) + (button_height * number_of_scripts_in_this_category))		'Margin x 3 because we need extra padding on the top, and half the extra on the bottom
				GroupBox left_col_x_pos, groupbox_y_pos, button_width + (groupbox_margin * 2), (groupbox_margin * 3) + (button_height * number_of_scripts_in_this_category), current_script_category_from_list
				number_of_scripts_in_this_category = 1
				groupbox_y_pos = left_col_y_pos
				left_col_y_pos = left_col_y_pos + (groupbox_margin * 2)
			else
				left_col_y_pos = left_col_y_pos + button_height
				number_of_scripts_in_this_category = number_of_scripts_in_this_category + 1
			end if			

		NEXT

		'>>> Placing new scripts on the list! This happens in the right-hand column of the dialog.
		
		right_col_y_pos = dialog_margin + (groupbox_margin * 2)
		
				
		'>>> Now we increment through the new scripts, and create buttons for them
		for i = 1 to num_of_new_scripts
			PushButton 215, right_col_y_pos, button_width, button_height, ucase(new_cs_scripts_array(i).category) & " - " & new_cs_scripts_array(i).script_name, button_placeholder
			button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100
			right_col_y_pos = right_col_y_pos + button_height
		NEXT

		GroupBox 210, dialog_margin, button_width + (dialog_margin * 2), 5 + (button_height * (UBound(new_cs_scripts_array) + 1)), "NEW SCRIPTS (LAST 60 DAYS)"
		
		
		' PushButton 210, dlg_height - 25, 60, 15, "Update Hotkeys", update_hotkeys_button						<<<<< SEE ISSUE #765
		PushButton 285, dlg_height - 25, 65, 15, "Update Favorites", update_favorites_button
		CancelButton 355, dlg_height - 25, 50, 15
	EndDialog

	'>>> Loading the favorites dialog
	DIALOG favorites_dialog
	'>>> Cancelling the script if ButtonPressed = 0
	IF ButtonPressed = 0 THEN stopscript
	
	
	
	'>>> Giving user has the option of updating their favorites menu.
	'>>> We should try to incorporate the chainloading function of the new script_end_procedure to bring the user back to their favorites.
	IF buttonpressed = update_favorites_button THEN
		call edit_favorites
		StopScript
	ElseIf buttonpressed = update_hotkeys_button then
		call edit_hotkeys
		StopScript
	End if
	
	'Determining the script that was selected, simply by subtracting 100 from the button_placeholder we'd previously defined. This corresponds with the array item selected.
	selected_script = ButtonPressed - 100
	
	'If it's a new script, it'll be larger than the text file array since it's displayed after, so this will create a variable for that too.
	selected_new_script = selected_script - (ubound(favorites_text_file_array) - 1)
	
	'This part takes the selected script integer and determines the file path for it
	if selected_script < ubound(favorites_text_file_array) then
		script_location = favorited_scripts_array(selected_script, 1)  '!!!! This works in conjunction with the button_placeholder that's used and incremented for each button. It won't work otherwise.
	else
		script_location = new_cs_scripts_array(selected_new_script).category & "/" & new_cs_scripts_array(selected_new_script).file_name
	end if
END FUNCTION
'======================================

'The script starts HERE!!!-------------------------------------------------------------------------------------------------------------------------------------

'>>> The gobbins of the script that the user sees and makes do.
'>>> Declaring the text file storing the user's favorite scripts list.
Dim oTxtFile
With (CreateObject("Scripting.FileSystemObject"))
	'>>> If the file exists, we will grab the list of the user's favorite scripts and run the favorites menu.
	If .FileExists(favorites_text_file_location) Then
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
		fav_scripts_command.Close
	ELSE
		'>>> ...otherwise, if the file does not exist, the script will require the user to select their favorite scripts.
		call edit_favorites
	END IF
END WITH

'>>> Calling the function that builds the favorites menu.
CALL favorite_menu(favorites_text_file_array, script_location)

''Figuring out where the script goes...
'if selected_script < ubound(favorites_text_file_array) then
'	script_location = favorited_scripts_array(selected_script, 1)
'end if
'
script_URL = script_repository & script_location


'>>> Running the script
If run_locally = true then
    Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
    Set fso_command = run_another_script_fso.OpenTextFile(script_URL)
    text_from_the_other_script = fso_command.ReadAll
    fso_command.Close
    Execute text_from_the_other_script
Else
    CALL run_from_GitHub(script_URL)
End if




'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< AGAIN, VERY TEMPORARY
'END IF
