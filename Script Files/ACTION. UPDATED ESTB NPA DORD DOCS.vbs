Enter file contents hereOption Explicit

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
			StopScript
END IF 

'THIS SCRIPT PRINTS THE CP AND NCP FINANCIAL STATEMENTS, WAIVERS, IMPORTANT STATEMENT OF RIGHTS, YOUR PRIVACY RIGHTS AND THE
'NCP NOTICE OF LIABILITY (F0100) FOR SENDING OUT ESTABLISHMENT DOCS ON A NPA OR DWP CASE. IT ALSO WRITES THE CAAD
'NOTE OF THE DOCUMENTS BEING SENT OUT AND ADDS A CAWT NOTE FOR RETURN IN 14 DAYS
'NOTE: NOT FOR A RELATIVE CARETAKER CASE AS THE CP WOULD NOT GET FINANCIAL DOCUMENTS  

'CONNECTING TO BLUEZONE 
EMConnect ""   

'CHECKS TO MAKE SURE WE ARE IN PRISM
CALL check_for_Prism (true)

'GOING TO CAAD SCREEN
Call navigate_to_Prism_Screen ("CAAD")

'SETTING THE CURSOR
EMSetCursor 21,18

'DIRECTING TO DORD
EMWriteScreen "DORD", 21,18

'ENTER
Transmit

'CLEARS THE DORD SCREEN
EMWriteScreen "C", 3,29

Transmit

'SETTING THE CURSOR
EMSetCursor 3,29

'ADDING THE DOC IN DORD
EMWriteScreen "A", 3,29

EMSetCursor 6,36
			
'FINANCIAL STATEMENT 
EMWriteScreen "F0021", 6,36

'ADDING THE FINANCIAL STATEMENT TO DORD
Transmit 

pf9

'PRINTING FINANCIAL STATEMENT
Transmit

pf9

'PRINTING FINANCIAL STATEMENT
Transmit 


EMSetCursor 3,29

'CLEARING DORD SCREEN FOR NEXT DOCUMENT
EMWriteScreen "C", 3,29

Transmit

'ADDING THE IMPORTANT STATEMENT OF RIGHTS IN DORD
EMWriteScreen "A", 3,29

EMSetCursor 6,36
		
'IMPORTANT STATEMENT OF RIGHTS
EMWriteScreen "F0022", 6,36

Transmit

pf9 'PRINTING IMPORTANT STATEMENT OF RIGHTS

Transmit

pf9 'PRINTING IMPORTANT STATEMENT OF RIGHTS

Transmit

EMSetCursor 3,29

'CLEARING THE DORD SCREEN FOR NEXT DOCUMENT
EMWriteScreen "C", 3,29

Transmit

'ADDING CP WAIVER DOC
EMWriteScreen "A", 3,29

EMSetCursor 6,36
		
'WAIVER
EMWriteScreen "F5000", 6,36

EMSetCursor 11,51

'CHANGING RECEIPIENT TO CP FOR WAIVER
EMWriteScreen "CPP", 11,51

Transmit

EMSetCursor 3,29

'MODIFYING DORD DOC SREEEN
EMWriteScreen "M", 3,29	   

'GOING INTO THE LABELS 			    
pf14

'SCROLLING DOWN 
pf8

'SETTING LINE TO MODIFY MONTHS ON WAIVER
EMSetCursor 13,5

'SELECTING LINE FOR MONTHS WAIVER IS VALID
EMWriteScreen "S", 13,5

Transmit

'ADDING 12 MONTHS FOR WAIVER BEING VALID
EMWriteScreen "12", 16,15

Transmit

pf3

pf9 'PRINTING CP WAIVER

Transmit

EMSetCursor 3,29

'CLEARING DORD SCREEN FOR NEXT DOCUMENT
EMWriteScreen "C", 3,29

Transmit

'ADDING THE DORD DOC
EmWriteScreen "A", 3,29

EMSetCursor 6,36

'ADDING THE NCP WAIVER DOC
EMWriteScreen "F5000", 6,36

EMSetCursor 11,51	
	
'CHANGING RECIPIENT TO NCP FOR THE WAIVER
EMWriteScreen "NCP", 11,51

Transmit

EMSetCursor 3,29

'MODIFYING DORD DOC SCREEN
EMWriteScreen "M", 3,29 

'GOING INTO THE LABELS
pf14
 
'SCROLLING DOWN
pf8

EMSetCursor 13,5

'SELECTING LINE FOR MONTHS WAIVER IS VALID
EMWriteScreen "S", 13,5

Transmit

'ADDING 12 MONTHS FOR WAIVER BEING VALID
EMWriteScreen "12", 16,15

Transmit

pf3

pf9 'PRINTING THE NCP WAIVER

Transmit	

EMSetCursor 3,29

'CLEARING THE DORD SCREEN
EMWriteScreen "C", 3,29

Transmit

'ADDING THE NCP AUTHORIZATION TO COLLECT SUPPORT DOC (F0100)
EMWriteScreen "A", 3,29  

EMSetCursor 6,36

'NCP NOTICE OF LIABILITY
EmwriteScreen "F0100", 6,36

Transmit

'MODIFYING THE AUTHORIZATION TO COLLECT TO ADD WORKER INFO
EMWriteScreen "M", 3,29

pf14 'GOING INTO THE LABELS

EMSetCursor 20,14

'GOING TO LABEL LINES
EMWriteScreen "U", 20,14

Transmit

EMSetCursor 7,5

'SELECTING THE LINE TO INCLUDE FINANCIAL STATEMENT ON THE NCP AUTHORIZATION TO COLLECT SUPPORT NOTICE
EMWriteScreen "S", 7,5

Transmit

'INCLUDING FINANCIAL STATEMENT LANGUAGE
EMwriteScreen "X", 16,15 

Transmit

pf3


'DIALOG THAT ENTERS THE WORKER NAME, TITLE, PHONE
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

'SELECTING LINE IN DORD DOC FOR WORKER NAME
EMWriteScreen "S", 9,5

Transmit 

'WRITING THE WORKER INFO INTO THE DORD DOC
EMWriteScreen "CSO_Name_Dialog", 16,15  'CSO name entered in dialog box

EMWriteScreen "CSO_Title_Dialog", 17,15 ' writes cso name into the dialog box

EMWriteScreen "CSO_Phone_Dialog", 18,15 ' writes cso phone into the dialog box

Transmit

EMSetCursor 10,05

'SELECTING LINE IN DORD DOC FOR WORKER TITLE
EMWriteScreen "S", 10,05

Transmit

EMSetCursor 11,5

'SELECTING LINE IN DORD DOC FOR WORKER PHONE
EMWriteScreen "S", 11,5

Transmit

pf3

pf9 'PRINTING THE NCP AUTHORIZATION TO COLLECT

'GOING TO CAAD TO WRITE THE NOTE
CALL navigate_to_PRISM_screen ("CAAD")  

pf5

'ADDING THE CAAD NOTE
EMWriteScreen "A", 3, 29

EMWriteScreen "FREE", 4, 54

EMSetCursor 16, 4

'WRITING THE CAAD NOTE
EMWriteScreen "SENT CP & NCP FINANCIAL STATEMENTS & WAIVERS, NCP AUTHORIZATION TO COLLECT_", 16,4

'EMWriteScreen "Authorization to collect sent to NCP", 17,4

Transmit

'GOING TO CAWT TO WRITE DUE DATE FOR RETURN OF FORMS
Call navigate_to_Prism_Screen ("CAWT")  

pf5

EMSetCursor 4,37

EMWriteScreen "Free", 4,37

EMSetCursor 10,4

'WRITING THE CAWT NOTE
EMWriteScreen "Did CP & NCP Return Fin Stmts & Waivers?", 10,4

EMSetCursor 17,52

'SETTING THE DUE DATE OUT FOR 14 DAYS
EMWriteScreen "14", 17,52

Transmit


StopScript


