Option Explicit
'Declare variables

DIM beta_agency
'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO 					'Declares variables to be good to option explicit users
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

'************************************************************************************************************************************

Dim Document_Generation_Dialog, Case_Number, View_Button, Print_Button, Document_Name_Listbox, ButtonPressed, objDoc

'Set up dialog box for document generation
BeginDialog Document_Generation_Dialog, 0, 0, 251, 140, "Document Generation"
  Text 5, 15, 50, 10, "Case Number:" 
  EditBox 55, 15, 95, 15, Case_Number  'The user enters the case number
  Text 5, 55, 115, 10, "Select a document to view or print:"
  DropListBox 125, 55, 95, 15, "<Select One>"+chr(9) +"Affidavit"+chr(9)+"CP Statement of Support"+chr(9)+"IW Notice"+chr(9)+"Nonpay Letter"+chr(9)+"QC Letter NPA"+chr(9)+"QC Letter PA", Document_Name_Listbox 'The user selects a document
  ButtonGroup ButtonPressed 'Three buttons: View, Print, and Cancel
    PushButton 85, 110, 50, 15, "&View", View_Button
    PushButton 140, 110, 50, 15, "&Print", Print_Button
    CancelButton 195, 110, 50, 15
EndDialog

EMConnect ""  'Connecting to bluezone

CALL check_for_PRISM(True) 'If not in PRISM, stop script

Do
Dialog Document_Generation_Dialog   'Show dialog box 
	If ButtonPressed = 0 then stopscript
	If Case_Number = "" then Msgbox "Please enter a case number."
	If Document_Name_Listbox = "<Select One>" then Msgbox "Please select a document."
Loop Until Case_Number <> "" and Document_Name_Listbox <> "<Select One>"

'If user selects cancel, stop the script
If ButtonPressed = 0 then stopscript

function ViewDocument (objDoc) 'function for viewing document
Msgbox "View " & Document_Name_Listbox & " for " & Case_Number
end function

function PrintDocument (objDoc) 'function for printing document
Msgbox "Print " & Document_Name_Listbox & " for " & Case_Number
end function

'If the user selects view, then view the document (call the viewing function)
If ButtonPressed = View_Button  then  
	ViewDocument(Document_Name_Listbox)
End If
'If the user selects print, then print the document to the user's default printer and to the virtual printer (call the printing function)
If ButtonPressed = Print_Button then 
	PrintDocument (Document_Name_Listbox)
End If
