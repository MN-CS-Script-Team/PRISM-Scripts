'Gathering stats+=====================
'name_of_scripts = "ACTIONS - ESTB DORD DOCS FOR PA CASE.vbs"
'start_time = timer



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
END IF  

'THE SCRIPT IS READY FOR USE

'This is an updated version of the ESTB NPA DORD DOCS that is used when starting a NEW ESTABLISH ACTION on a non public assistance or DWP case that prints the
'the financial statements, waivers, Important Statement of Rights and NCP Authorization to Collect Support. It would NOT be used on a RELATIVE CARETAKER case as
'the CP is not required to complete financial docs on that type of case.

'Connecting to BZ  'This is a script for a PA case to print the fin docs & waivers
EMConnect ""   

'Checks to make sure we are in Prism
CALL check_for_Prism (true)

'Directing to DORD screen
Call navigate_to_Prism_Screen ("DORD")

'Clears the screen to add the doc
EMWriteScreen "C", 3,29

Transmit

EMWriteScreen "A", 3,29

EMSetCursor 6,36
		
'adding the financial statement to DORD
EMWriteScreen "F0021", 6,36

'Printing financial statement x 2
Transmit

pf9

transmit

pf9

transmit

EMSetCursor 3,29

EMWriteScreen "C", 3,29

transmit

'Adding the Important Statement of Rights
EMWriteScreen "A", 3,29

EMSetCursor 6,36			

EMWriteScreen "F0022", 6,36

'Printing the Important Statement of Rights x 2
Transmit

pf9

transmit

pf9

transmit

EMSetCursor 3,29

'Clearing screen for next doc
EMWriteScreen "C", 3,29

transmit

'Adding CP Waiver
EMWriteScreen "A", 3,29

EMSetCursor 6,36		

EMWriteScreen "F5000", 6,36

EMSetCursor 11,51

'Changing recipient to CP in DORD
EMWriteScreen "CPP", 11,51

transmit

EMSetCursor 3,29

'Modifying label in DORD
EMWriteScreen "M", 3,29	   
				    
pf14

pf8

EMSetCursor 13,5

'Selecting label line
EMWriteScreen "S", 13,5

transmit

'updating Waiver to say 12 months valid
EMWriteScreen "12", 16,15

Transmit

pf3

'printing CP Waiver
pf9

transmit

EMSetCursor 3,29

'Clearing screen for next doc
EMWriteScreen "C", 3,29

transmit

'Adding Waiver to DORD
EmWriteScreen "A", 3,29

EMSetCursor 6,36

EMWriteScreen "F5000", 6,36

EMSetCursor 11,51		

'Changing recipient to NCP on Waiver
EMWriteScreen "NCP", 11,51

transmit

EMSetCursor 3,29

'Modifying label on DORD doc
EMWriteScreen "M", 3,29

pf14

pf8

EMSetCursor 13,5

'Selecting label line
EMWriteScreen "S", 13,5

transmit

'Modifying label to say Waiver valid for 12 months
EMWriteScreen "12", 16,15

transmit

pf3

'Printing NCP Waiver
pf9

transmit	

EMSetCursor 3,29

'Clearing DORD screen
EMWriteScreen "C", 3,29

transmit

'Adding the NCP Authorization to Collect Support 
EMWriteScreen "A", 3,29   

EMSetCursor 6,36

EmwriteScreen "F0100", 6,36

transmit

EMWriteScreen "M", 3,29

pf14

EMSetCursor 20,14

EMWriteScreen "U", 20,14

transmit

'Selecting label line to include financial statement language on DORD doc
EMSetCursor 7,5

EMWriteScreen "S", 7,5

transmit

'Selecting the "Include Financial Statement" line
EMwriteScreen "X", 16,15 

transmit


'The Dialog to add worker information in the labels

DIM pa_dord_docs_dialog, worker_name_dialog, worker_title_dialog, worker_phone_dialog, ButtonPressed, write_variable_in_DORD
'INSERTED THE WORKER INFORMATION NEW DIALOG HERE

BeginDialog npa_dord_docs_dialog, 0, 0, 191, 135, "Worker Information Dialog"
  Text 10, 10, 50, 10, "Worker Name:"
  Text 10, 35, 45, 10, "Worker Title:"
  Text 10, 60, 55, 10, "Worker Phone:"
  EditBox 60, 5, 115, 15, worker_name_dialog
  EditBox 55, 30, 120, 15, worker_title_dialog
  EditBox 65, 55, 110, 15, worker_phone_dialog
  ButtonGroup ButtonPressed
    OkButton 10, 90, 50, 15
    CancelButton 10, 110, 50, 15
EndDialog

'This makes the dialog run
Dialog npa_dord_docs_dialog  

IF ButtonPressed = 0 THEN StopScript

EMSetCursor 9,5

EMWriteScreen "S", 9,5

transmit     'This next part below is needed to make the dialog box run in the script when entering the info

'Below writes the worker information typed into the DORD doc
EMWriteScreen (worker_name_dialog), 16,15 

transmit

EMSetCursor 10,5
EMWriteScreen "S", 10,5

transmit

EMWriteScreen (worker_title_dialog), 16,15  

transmit

EMSetCursor 11,5

EMWriteScreen "S", 11,5

transmit

EMWriteScreen (worker_phone_dialog), 16,15  

transmit

pf3

pf9

CALL navigate_to_PRISM_screen ("CAAD")  

PF5

EMSetCursor 4,54    

EMWriteScreen "FREE", 4,54

EMSetCursor 16,4

EMWriteScreen "Sent CP and NCP Financial Statements and Waivers", 16,4

transmit

'Going to CAWT to write the tracking of the forms
Call navigate_to_Prism_Screen ("CAWT")

pf5

EMSetCursor 3,3

EMWriteScreen "A", 3,3

EMSetCursor 4,37   

EMWriteScreen "Free", 4,37

EMSetCursor 10,4

'Writing the CAWT note out to watch for return of forms in 14 days
EMWriteScreen "Did CP & NCP Return Financial Statements and Waivers?",10,4

EMSetCursor 17,52

'Setting out the CAWT for 14 days
EMWriteScreen "14", 17,52

transmit

pf3

script_end_procedure("")
