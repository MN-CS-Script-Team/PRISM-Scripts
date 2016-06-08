'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - EST DORD NPA DOCS.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED

DIM beta_agency

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
END IF   'Remember the pf9s are not set to print, they are in green, remove the ' when ready to use

'Connecting to BZ  'This is a script for a NPA or DWP case to print the fin docs
EMConnect ""   

'Checks to make sure we are in Prism
CALL check_for_Prism (true)

'Gets to CAAD
Call navigate_to_Prism_Screen ("CAAD")

'Set cursor on the CAAD line
EMSetCursor 21,18

'Directing to Dord
EMWriteScreen "DORD", 21,18

'hits the Enter Key
Transmit

'Sets to add mode
EMWriteScreen "C", 3,29

Transmit

EMSetCursor 3,29

EMWriteScreen "A", 3,29

EMSetCursor 6,36			'Printing the Financial Statement x 2

EMWriteScreen "F0021", 6,36

Transmit

pf9

Transmit

pf9

Transmit

EMSetCursor 3,29

EMWriteScreen "C", 3,29

Transmit

EMWriteScreen "A", 3,29

EMSetCursor 6,36			'Printing the Important Statement of Rights x 2

EMWriteScreen "F0022", 6,36

Transmit

pf9

Transmit

pf9

Transmit

EMSetCursor 3,29

EMWriteScreen "C", 3,29

Transmit

EMWriteScreen "A", 3,29

EMSetCursor 6,36			'Printing the CP Waiver

EMWriteScreen "F5000", 6,36

EMSetCursor 11,51

EMWriteScreen "CPP", 11,51

Transmit

EMSetCursor 3,29

EMWriteScreen "M", 3,29	   
				    
pf14

pf8

EMSetCursor 13,5

EMWriteScreen "S", 13,5

Transmit

EMWriteScreen "12", 16,15

Transmit

pf3

pf9

Transmit

EMSetCursor 3,29

EMWriteScreen "C", 3,29

Transmit

EmWriteScreen "A", 3,29

EMSetCursor 6,36

EMWriteScreen "F5000", 6,36

EMSetCursor 11,51		'Printing the NCP Waiver

EMWriteScreen "NCP", 11,51

Transmit

EMSetCursor 3,29

EMWriteScreen "M", 3,29

pf14

pf8

EMSetCursor 13,5

EMWriteScreen "S", 13,5

Transmit

EMWriteScreen "12", 16,15

Transmit

pf3

pf9

Transmit	

EMSetCursor 3,29

EMWriteScreen "C", 3,29

Transmit

EMWriteScreen "A", 3,29   'Printing the NCP Authorization to Collect Support

EMSetCursor 6,36

EmwriteScreen "F0100", 6,36

Transmit

EMWriteScreen "M", 3,29

pf14

EMSetCursor 20,14

EMWriteScreen "U", 20,14

Transmit

EMSetCursor 7,5

EMWriteScreen "S", 7,5

Transmit

EMwriteScreen "X", 16,15 

Transmit

DIM Dialog1, CSO_Name_Dialog, CSO_Title_Dialog, CSO_Phone_Dialog, ButtonPressed, write_variable_in_DORD


BeginDialog Dialog1, 0, 0, 191, 135, "CSO Information"
  Text 10, 15, 35, 10, "CSO Name"
  EditBox 55, 10, 105, 15, CSO_Name_Dialog
  Text 10, 40, 40, 10, "CSO Title"
  EditBox 55, 35, 105, 15, CSO_Title_Dialog
  Text 10, 65, 50, 10, "CSO Phone No"
  EditBox 65, 60, 65, 15, CSO_Phone_Dialog
  ButtonGroup ButtonPressed
    OkButton 65, 85, 50, 15
    CancelButton 65, 105, 50, 15
EndDialog

Dialog Dialog1

IF ButtonPressed = 0 THEN StopScript

EMSetCursor 9,5

EMWriteScreen "S", 9,5

Transmit  'The info below is needed to make the dialog box run in the script when entering CSO info

EMWriteScreen (CSO_Name_Dialog), 16,15  'CSO name entered in dialog box

Transmit

EMSetCursor 10,05

EMWriteScreen "S", 10,05

Transmit

EMWriteScreen (CSO_Title_Dialog), 16,15   'CSO title entered in dialog box

Transmit

EMSetCursor 11,5

EMWriteScreen "S", 11,5

Transmit

EMWriteScreen (CSO_Phone_Dialog), 16,15  'CSO phone entered in dialog box

Transmit

pf9

pf3

CALL navigate_to_PRISM_screen ("CAAD")  'Writing the CAAD note of docs sent out

pf5

EMWriteScreen "A", 3, 29

EMWriteScreen "FREE", 4, 54

EMSetCursor 16, 4

EMWriteScreen "Financial Statements & Waivers sent to parties", 16,4

EMWriteScreen "Authorization to collect sent to NCP", 17,4

Transmit

pf3

Call navigate_to_Prism_Screen ("CAWT")  'Adding the note to CAWT with due date of forms

pf5

EMSetCursor 4,37

EMWriteScreen "Free", 4,37

EMSetCursor 10,4

EMWriteScreen "CP & NCP Return Fin Stmts & Waivers?",10,4

EMSetCursor 17,52

EMWriteScreen "14", 17,52

Transmit

pf3

StopScript

