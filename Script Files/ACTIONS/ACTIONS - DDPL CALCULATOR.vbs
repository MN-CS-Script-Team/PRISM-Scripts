'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - DDPL CALCULATOR.vbs"
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
END IF	

'Declared
DIM ddpl_calculator, PRISM_MCI_number, PRISM_begin_date, PRISM_end_date, buttonpressed, row, direct_deposit_issued_date, end_of_data_check, direct_deposit_amount, end_date, total_amount_issued, string_for_msgbox

'DDPL Dialog Box
BeginDialog ddpl_calculator, 0, 0, 191, 105, "DDPL Calculator"
  ButtonGroup ButtonPressed
    OkButton 80, 80, 50, 15
    CancelButton 135, 80, 50, 15
  Text 15, 10, 65, 10, "PRISM MCI Number"
  EditBox 95, 5, 60, 15, PRISM_MCI_number
  Text 35, 30, 50, 10, "Start Date"
  EditBox 95, 25, 50, 15, PRISM_begin_date
  Text 35, 50, 50, 10, "End Date"
  EditBox 95, 45, 50, 15, PRISM_end_date
EndDialog

Dialog ddpl_calculator

IF ButtonPressed = cancel THEN StopScript

EMConnect ""

CALL check_for_prism(TRUE)

CALL navigate_to_prism_screen ("DDPL")

EMWriteScreen PRISM_MCI_number, 20, 007

EMSendKey "<enter>"

EMWaitReady 0,0

EMWriteScreen PRISM_begin_date, 20, 038

EMWriteScreen PRISM_end_date, 20, 067

transmit

row = 8

total_amount_issued = 0

Do 
	
EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list 
	If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have 
EMReadScreen direct_deposit_issued_date, 9, row, 11 				'Reading the issue date 
EMReadScreen direct_deposit_amount, 10, row, 33 				'Reading amount issued 

total_amount_issued = total_amount_issued + abs(direct_deposit_amount) 	'Totals amount issued 

row = row + 1 										'Increases the row variable by one, to check the next row 

EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list 
    If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have 

    If row = 19 then 									'Resets row and PF8s 
        PF8 
        row = 8 
    End if 
Loop until end_of_data_check = "*** End of Data ***" 

string_for_msgbox = " Total payments issued for the period of " & PRISM_begin_date & " through " & PRISM_end_date & " is $" & total_amount_issued 

MsgBox string_for_msgbox 
script_end_procedure("")
