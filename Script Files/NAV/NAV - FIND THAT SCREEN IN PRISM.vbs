'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - FIND THAT SCREEN IN PRISM.vbs"
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


BeginDialog Find_that_screen_in_PRISM_dialog, 0, 0, 161, 185, "Find that Screen in PRISM"
  ButtonGroup ButtonPressed
    OkButton 45, 160, 50, 15
    CancelButton 100, 160, 50, 15
    PushButton 115, 5, 35, 10, "ACSD", ACSD_button
    PushButton 115, 20, 35, 10, "CPRE", CPRE_button
    PushButton 115, 35, 35, 10, "GCSC", GCSC_button
    PushButton 115, 50, 35, 10, "NCLD", NCLD_button
    PushButton 115, 65, 35, 10, "NCLL", NCLL_button
    PushButton 115, 80, 35, 10, "NCSL", NCSL_button
    PushButton 115, 95, 35, 10, "PALI", PALI_button
    PushButton 115, 110, 35, 10, "REID", REID_button
    PushButton 115, 125, 35, 10, "SEPD", SEPD_button
    PushButton 115, 140, 35, 10, "WEDL", WEDL_button
  Text 10, 5, 70, 10, "Account status detail"
  Text 10, 20, 70, 10, "CP/NCP Relationship"
  Text 10, 35, 70, 10, "Good Cause Screen"
  Text 10, 50, 80, 10, "NCP license data detail"
  Text 10, 65, 70, 10, "NCP license data list"
  Text 10, 80, 70, 10, "NCP alias detail"
  Text 10, 95, 70, 10, "Payment listing"
  Text 10, 110, 90, 10, "Re-employment ins detail"
  Text 10, 125, 90, 10, "Service of Process detail"
  Text 10, 140, 95, 10, "On-line employer reporting"
EndDialog

'The Script--------------------------------------------------------------------------------------------------

Dialog Find_that_screen_in_PRISM_dialog

'Connect to BlueZone
EMConnect ""

CALL check_for_PRISM(true)			'this checks whether in PRISM or timed out

'Not adding case number function as PRISM will go to screen without case number or PRISM populates with last case number used.

'Naviagting to any of the buttons chosen
If buttonpressed = ACSD_button then call navigate_to_PRISM_screen("ACSD")   
If buttonpressed = CPRE_button then call navigate_to_PRISM_screen("CPRE")
If buttonpressed = GCSC_button then call navigate_to_PRISM_screen("GCSC")   
If buttonpressed = NCLD_button then call navigate_to_PRISM_screen("NCLD")
If buttonpressed = NCLL_button then call navigate_to_PRISM_screen("NCLL")   
If buttonpressed = NCSL_button then call navigate_to_PRISM_screen("NCSL")
If buttonpressed = PALI_button then call navigate_to_PRISM_screen("PALI")   
If buttonpressed = REID_button then call navigate_to_PRISM_screen("REID")
If buttonpressed = SEPD_button then call navigate_to_PRISM_screen("SEPD")   
If buttonpressed = WEDL_button then call navigate_to_PRISM_screen("WEDL")

script_end_procedure("")

