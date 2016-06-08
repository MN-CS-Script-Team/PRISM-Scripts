'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "BULK - F0320 Scrubber.vbs" 
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

' >>>>> THE SCRIPT <<<<<
EMConnect ""

'>>>>> GOING TO USWT <<<<<
CALL navigate_to_Prism_screen("USWT")

' >>>>> SELECTING THE SPECIFIC WORKLIST TYPE <<<<<
EMWriteScreen "F0320", 20, 30
transmit

USWT_row = 7
COUNT = 0
SCROLL = 0
' >>>>> STARTING THE DO LOOP. THE SCRIPT NEEDS TO HANDLE THESE CASES ONE AT A TIME <<<<<
DO
	EMReadScreen USWT_type, 5, USWT_row, 45
	IF USWT_type = "F0320" THEN
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		EMWriteScreen "d", USWT_row, 4
		transmit
		'Selecting the worklist brings the user to NCP's PAPL screen	
		purge = false 'Reset the purge variable
	END IF
	CALL navigate_to_PRISM_screen ("SUOD")	
	EMWriteScreen "B", 3, 29
	transmit
	
	EMReadScreen order_type, 3, 10, 22
	IF order_type <> "CTM" THEN 
		EMSetCursor 10, 8
		Transmit
		PF11
		EMReadScreen med_code, 3, 12, 74
		If med_code = "PAO" THEN
			purge = true
		END IF
	END IF

		
	CALL navigate_to_PRISM_screen ("CAWT")
	EMWriteScreen "F0320", 20, 29
	EMWriteScreen USWT_case_number, 20, 8
	transmit

		' >>>>> IF THE WORKLIST ITEM IS ELIGIBLE TO BE PURGED, THE SCRIPT PURGES...
	IF purge = TRUE THEN 
	CAWT_row = 8
		DO 
			EMReadScreen CAWD_type, 5, CAWT_row, 8
			IF cawd_type = "F0320" then	
				EMWriteScreen "P", CAWT_row, 4
				transmit
				transmit
				Count = Count + 1
				PF3
			END IF
			cawt_row = cawt_row + 1
		LOOP UNTIL cawd_type <> "F0320"
	END IF
		'  ...  IF THE WORKLIST ITME IS NOT ELIGIBLE TO BE PURGED, THE SCRIPT INCREASES USWT_ROW + 1 <<<<<
	CALL navigate_to_PRISM_screen ("USWT")

	EMWriteScreen "F0320", 20, 30
	transmit
	IF SCROLL > 0 THEN
		FOR I = 0 TO SCROLL
			PF8
		NEXT
	END IF
	USWT_row = USWT_row + 1
	IF USWT_row = 19 THEN 
		PF8
		USWT_row = 7
		SCROLL = SCROLL + 1
	END IF
		
	
LOOP UNTIL USWT_type <> "F0320"

script_end_procedure("Success!  " & Count & " worklists purged!")
