'Option Explicit

'STATS GATHERING ---------------------------
name_of_script = "ACTIONS - F5002.vbs"
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


'Connects to Bluezone
EMConnect ""

EMReadScreen screen_name, 19, 2, 29
If screen_name <> "CODO Request Detail" THEN
	MSGBOX "Re-run this script after you have added your document in CORD and have completed the legal heading, legal tracking, GUWD worksheet, and partys' address and employment selections.  Script is now ending."
	stopscript
END IF


	DO
		EMReadScreen heading, 11, 10, 22
		IF heading <> "Description" THEN
			PF11	
		END IF
	LOOP UNTIL heading = "Description"

PF8
PF8
PF8

'Adding Form 11.2 paragraph - 351
EMWriteScreen "N", 15, 3
save_cord_doc
EMWriteScreen "D", 16, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph ("Form 11.2 has been prepared with attachments in support of this motion.")
transmit
PF3
save_cord_doc

PF8
'Selecting the paragraphs that indicate the joint child resides with CP - 420 & 430
EMWriteScreen "S", 13, 3
EMWriteScreen "S", 14, 3
save_cord_doc


PF8
PF8
PF8

'Deselecting household expenses paragraphs - 740 & 770
EMWriteScreen " ", 15, 3
EMWriteScreen " ", 18, 3
save_cord_doc

'Adding options for General Assistance and the special Medical Assistance language - 801 & 802
PF8
EMWriteScreen "B", 11, 3
save_cord_doc
EMWriteScreen "D", 12, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph ("General Assistance.")
transmit
PF3
save_cord_doc
EMWriteScreen "B", 12, 3
save_cord_doc
EMWriteScreen "D", 13, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Due to automated data sharing issues between the state medical assistance computer system (METS)" _
	& " and the state child support computer system (PRISM), the County Child Support Program does not have" _
	& " access to Medical Assistance information on some cases, including whether Medical Assistance is in" _
	& " place or how much was expended.")
transmit
PF3
save_cord_doc

'Deselecting General Assistance - 801
EMWriteScreen " ", 12, 3  
save_cord_doc

PF8

'Adding option for calculating NCP's DEED info - 951
EMWriteScreen "B", 18, 3
save_cord_doc
EMWriteScreen "D", 19, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Calculating the total of quarterly wages reported to the Minnesota Department of Employment"_
	& " and Economic Development from __/__/__ to __/__/___, and dividing by ___ months.")
transmit
PF3
save_cord_doc

PF8
PF8
PF8

'Selecting paragraph for NCP's nonjoint children in household other than NCP's - 1250
EMWriteScreen "S", 19, 3
save_cord_doc

PF8

'Selecting paragraph for NCP's nonjoint children in NCP's household - 1320
EMWriteScreen "S", 16, 3
save_cord_doc

'Adding new options for CP's public assistance programs that do not indicate the amount of benefit received - 1340-1344
EMWriteScreen "B", 18, 3
save_cord_doc
EMWriteScreen "D", 19, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Minnesota Family Investment Program (MFIP) cash assistance, creating an assignment under"_
	& " Minnesota and Federal Law.")
transmit
PF3
save_cord_doc


EMWriteScreen "B", 19, 3
save_cord_doc
EMWriteScreen "D", 20, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Medical Assistance, creating an assignment under Minnesota and Federal Law.")
transmit
PF3
save_cord_doc


EMWriteScreen "B", 20, 3
save_cord_doc


PF8
EMWriteScreen "D", 11, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Child Care Assistance, creating an assignment under Minnesota and Federal Law.")
transmit
PF3
save_cord_doc

EMWriteScreen "B", 11, 3
save_cord_doc
EMWriteScreen "D", 12, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Due to automated data sharing issues between the state medical assistance computer system (METS)"_
	& " and the state child support computer system (PRISM), the County Child Support Program does not have"_
	& " access to Medical Assistance information on some cases, including whether Medical Assistance is in"_
	& " place or how much was expended.")
transmit
PF3
save_cord_doc

PF8

'Adding option for calculating CP's DEED info - 1491
EMWriteScreen "B", 17, 3
save_cord_doc
EMWriteScreen "D", 18, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph( "Calculating the total of quarterly wages reported to the Minnesota Department of Employment"_
	& " and Economic Development from __/__/__ to __/__/___, and dividing by ___ months.")
transmit
PF3
save_cord_doc


PF8
PF8
PF8

'Selecting paragraph for CP's nonjoint children in household other than CP's - 1790
EMWriteScreen "S", 18, 3
save_cord_doc

PF8

'Selecting paragraph for CP's nonjoint children in CP's household - 1860
EMWriteScreen "S", 15, 3
save_cord_doc

'Deselect paragraph for CP's PICS (Potential Income) - 1890
EMWriteScreen " ", 18, 3
save_cord_doc

PF8

'Select paragraph for NCP's basic support obligation - 1930
EMWriteScreen "S", 12, 3
save_cord_doc

'De-select paragraph for NCP's basic support obligation after parenting time adjustment - 1940
EMWriteScreen " ", 13, 3
save_cord_doc

PF8
PF8
PF8
PF8
PF8
PF8

'Select paragraph for children receive public Medical Assistance coverage - 2550
EMWriteScreen "S", 14, 3
save_cord_doc
EMWriteScreen "D", 14, 3
transmit
EMWriteScreen "D", 10, 4
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("0.00")
transmit
PF3
PF3
save_cord_doc

'Select paragraph for NCP's medical support obligation - 2580
EMWriteScreen "S", 17, 3
save_cord_doc

PF8

'Deselect Pursuant to MN Statutes, NCP must - 2630
EMWriteScreen " ", 12, 3
save_cord_doc



'Add new paragraph reserving medical support - 2661
EMWriteScreen "N", 15, 3
save_cord_doc

EMWriteScreen "D", 16, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("Due to automated data sharing issues between the state medical assistance computer system (METS)"_
	& " and the state child support computer system (PRISM), the County Child Support Program does not have"_
	& " access to Medical Assistance information on some cases, including whether Medical Assistance is in"_
	& " place or how much was expended.  Therefore, the county requests that the issue of medical support be"_
	& " reserved until medical public assistance information becomes available.  The county does not seek any"_
	& " reimbursement for medical assistance until further motion.")
transmit
PF3
save_cord_doc

'Deselect Pursuant to MN Statutes, CP must - 2660
EMWriteScreen " ", 15, 3
save_cord_doc

PF8

'Adding new paragraph about minimal unreimbursed/uninsured expenses while MA is in place - 2751
EMWriteScreen "N", 15, 3
save_cord_doc
EMWriteScreen "D", 16, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph( "There are minimal uninsured or unreimbursed health-related expenses while public coverage is"_
	& " in place.")
transmit
PF3
save_cord_doc

PF8
PF8
PF8

'Adding new paragraph for the State to waive public assistance expended (no amount) - 3031
EMWriteScreen "N", 14, 3
save_cord_doc

EMWriteScreen "D", 15, 3
transmit
EMWriteScreen "M", 8, 33
EMSetCursor 11, 6
write_variable_to_CORD_paragraph("The State of Minnesota waives the past support that may be due and owing up to the commencement"_
	& " date of the ongoing support obligations.")
transmit
PF3
save_cord_doc


PF8
PF8
PF8
PF8
PF8

'Deselecting the paragraph that NCP meets the minimum basic support requirements - 
EMWriteScreen " ", 19, 3
save_cord_doc


script_end_procedure("Script is now ending.  Your CORD document is ready for review/editing.")

