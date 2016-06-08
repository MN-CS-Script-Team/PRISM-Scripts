'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - PRISM SCREEN FINDER FOR FINANCIAL WORKERS.vbs"
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


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_screen_finder_dialog, 0, 0, 261, 135, "PRISM screen finder"
  ButtonGroup ButtonPressed
    CancelButton 210, 120, 50, 15
    PushButton 140, 70, 45, 10, "DDPL", DDPL_button
    PushButton 140, 40, 45, 10, "CAAD", CAAD_button
    PushButton 140, 55, 45, 10, "CAFS", CAFS_button
    PushButton 140, 85, 45, 10, "GCSC", GCSC_button
    PushButton 140, 115, 45, 10, "PESE", PESE_button
    PushButton 140, 100, 45, 10, "PALC", PALC_button
  Text 35, 70, 90, 10, "Direct deposit listing:"
  Text 35, 40, 65, 10, "Case notes:"
  Text 35, 55, 100, 10, "Case financial summary:"
  Text 35, 85, 100, 10, "Good cause/safety concerns:"
  Text 35, 100, 100, 10, "Case payment allocations:"
  Text 35, 115, 65, 10, "Person search:"
  Text 10, 0, 250, 25, "Press a button below to navigate to PRISM screens.  Then press F1 in the case number or MCI number field to select the participant or case information you are looking for."
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connect to BlueZone
EMConnect ""

CALL check_for_PRISM(FALSE)

Dialog PRISM_screen_finder_dialog

'Now it'll navigate to any of the screens chosen
If buttonpressed = DDPL_button then call navigate_to_PRISM_screen("DDPL")
If buttonpressed = CAAD_button then call navigate_to_PRISM_screen("CAAD")
If buttonpressed = CAFS_button then call navigate_to_PRISM_screen("CAFS")
If buttonpressed = GCSC_button then call navigate_to_PRISM_screen("GCSC")
If buttonpressed = PESE_button then call navigate_to_PRISM_screen("PESE")
If buttonpressed = PALC_button then call navigate_to_PRISM_screen("PALC")

script_end_procedure("")
