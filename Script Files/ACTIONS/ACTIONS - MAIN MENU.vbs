'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
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
			"Robert will work to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'Loading all scripts
CALL run_from_GitHub("https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/ALL%20SCRIPTS.vbs")

DIM ButtonPressed, button_placeholder
DIM SIR_instructions_button
DIM Dialog1

Function declare_main_menu(menu_type, script_array)
	BeginDialog Dialog1, 0, 0, 516, 340, menu_type & " Scripts"
	  ButtonGroup ButtonPressed
	 	'This starts here, but it shouldn't end here :)
		vert_button_position = 30
		button_placeholder = 100
		FOR current_script = 0 to ubound(script_array)
			IF InStr(script_array(current_script).script_type, menu_type) <> 0 THEN
				IF InStr(script_array(current_script).agencies_that_use, UCASE(replace(county_name, " County", ""))) <> 0 THEN 
					'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
					'FUNCTION		HORIZ. ITEM POSITION								VERT. ITEM POSITION		ITEM WIDTH									ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
					PushButton 		5, 													vert_button_position, 	script_array(current_script).button_size, 	10, 			script_array(current_script).script_name, 			button_placeholder
					Text 			script_array(current_script).button_size + 10, 		vert_button_position, 	500, 										10, 			"--- " & script_array(current_script).description
					'----------
					vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
					'----------
					script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				END IF
			END IF
			button_placeholder = button_placeholder + 1
		NEXT
		PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		CancelButton 460, 320, 50, 15
	EndDialog
End function

DO
	CALL declare_main_menu("ACTIONS", cs_scripts_array)
	Dialog
	IF ButtonPressed = 0 THEN script_end_procedure("")
	IF ButtonPressed = SIR_instructions_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/Shared%20Documents/Forms/All%20ACTIONS%20Scripts.aspx")
LOOP UNTIL ButtonPressed <> SIR_instructions_button

'Determining the script selected from the value of ButtonPressed
'Since we start at 100 and then go up, we will simply subtract 100 when determining the position in the array
script_picked = ButtonPressed - 100

'Running the selected script
CALL run_from_GitHub(script_repository & cs_scripts_array(script_picked).script_type & "/" & cs_scripts_array(script_picked).file_name)


