
'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - EMANCIPATION WORKLIST SCRUBBER.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block 


'this is a function document
DIM beta_agency 'remember to add


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

'--------------------------------------------------------

'connecting to bluezone
EMConnect ""

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

count = 0 
'------------------------------------------------------------
'Declaring Variables used in the loop
DIM M0935Str, M0935_Confirm, count
DIM Row, Child_Actv, Child_DOB, Child_Age, Child_MCI
DIM SUOD_Type, Child_Row, Child_Col, Emanc_Code

Do 
	CALL navigate_to_PRISM_screen("USWT")
	EMWriteScreen "M0935", 20, 30
	transmit

	'CONFIRMING WORKLIST IS M0935
	EMReadScreen M0935Str, 5, 7, 45
	If M0935Str <> "M0935" Then Exit Do		
	EMWriteScreen "D", 7, 4
	transmit
	

	'CONFIRMING WORKLIST HASN'T ALREADY BEEN REVIEWED
	EMReadScreen M0935_Confirm, 3, 10, 4	
	If M0935_Confirm <> "___" Then Exit Do
	CALL navigate_to_PRISM_screen("CHDE")	
	EMWriteScreen "B", 3, 29
	transmit

	'BEGINNING LOOP TO FIND CHILD
	Row = 8
	Do
		EMReadScreen Child_Actv, 1, Row, 35
		If Child_Actv = " " Then 
			'MsgBox "Unable to find child with an 18th birthday within the next 3 months! Please process worklist manually! Script Ended.", VBExclamation
			'StopScript
			script_end_procedure("Unable to find child with an 18th birthday within the next 3 months! Please process worklist manually! Script Ended.")

		ElseIf Child_Actv = "Y" Then
			EMReadScreen Child_DOB, 8, Row, 57
			'CONFIRMING CHILD'S 18TH BIRTHDAY WILL BE IN THE NEXT 3 MOS
			'BY CALCULATING CHILD'S DOB FROM TODAY'S DATE (MUST BE BETWEEN 213 AND 217 MONTHS)
			Child_Age = DateDiff("m", Child_DOB, Date)
			If (Child_Age >= 213) And (Child_Age <= 217) Then	
				EMReadScreen Child_MCI, 10, Row, 67
				Exit Do
			End If
		End If
	Row = Row + 1
	Loop

	'BEGINNING LOOP TO FIND COURT ORDER EMANCIPATION LANGUAGE
	CALL navigate_to_PRISM_screen("SUOL")
	transmit
	Row = 10
	Do
		'IF UNABLE TO FIND EMANCIPATION LANGUAGE ON ANY ORDER, UPDATING WORKLIST WITH NOTE
		EMReadScreen SUOD_Type, 3, Row, 22
		If SUOD_Type = "   " Then
			CALL navigate_to_PRISM_screen("USWT")
			
			EMWriteScreen "M0935", 20, 30
			transmit
			EMWriteScreen "M", 7, 4
			transmit
			EMWriteScreen "~.~REVIEWED BY M0935 WORKLIST ON " & Date, 10, 4
			EMWriteScreen "ORDER DOES NOT ADDRESS EMANCIPATION - FURTHER REVIEW NEEDED", 11, 4
			EMWriteScreen "          ", 17, 21
			EmWriteScreen "1", 17, 52
			transmit
			EMSendKey "<PF3>"
			EMWaitReady 10, 250
			Exit Do
		ElseIf SUOD_Type <> "   " Then 
			EMSetCursor Row, 72
			transmit
			PF11

			'LOOKING FOR CHILD'S MCI TO CONFIRM EMANCIPATION LANGUAGE
			Child_Row = 1
			Child_Col = 1
			EMSearch Child_MCI, Child_Row, Child_Col

			'IF CHILD'S EMANCIPATION LANGUAGE IS GR, DORD DOC F0300 AND F0302 ARE GENERATED
			If Child_Row > 0 Then
				EMreadScreen Emanc_Code, 2, Child_Row, 66
				If Emanc_Code = "GR" Then
					CALL navigate_to_PRISM_screen("DORD")
					transmit
					EMWriteScreen "C", 3, 29
					transmit
					EMWriteScreen "A", 3, 29
					EMWriteScreen "F0300", 6, 36
					transmit
					Child_Row = 1
					Child_Col = 1
					EMSearch Child_MCI, Child_Row, Child_Col
					EMSetCursor Child_Row, Child_Col
					transmit

					EMWriteScreen "C", 3, 29
					transmit
					EMWriteScreen "A", 3, 29
					EMWriteScreen "F0302", 6, 36
					transmit
					Child_Row = 1
					Child_Col = 1
					EMSearch Child_MCI, Child_Row, Child_Col
					EMSetCursor Child_Row, Child_Col
					transmit

					'PURGING WORKLIST (DOCS TO PRINT OVERNIGHT)
					CALL navigate_to_PRISM_screen("USWT")
					transmit
					EMWriteScreen "M0935", 20, 30
					transmit
					EMWriteScreen "P", 7, 4
					transmit
					transmit
					count = count + 1
					Exit Do

				'IF CHILD'S EMANCIPATION LANGUAGE IS NOT GR, UPDATING WORKLIST WITH NOTE
				ElseIf Emanc_Code <> "GR" AND Emanc_Code <> "__" Then
					CALL navigate_to_PRISM_screen("USWT")
					transmit
					EMWriteScreen "M0935", 20, 30
					transmit
					EMWriteScreen "M", 7, 4
					transmit
					EMWriteScreen "~.~REVIEWED BY M0935 WORKLIST ON " & Date, 10, 4
					EMWriteScreen "ORDER DOES NOT HAVE STANDARD GRADUATION LANGUAGE - FURTHER REVIEW NEEDED", 11, 4
					EMWriteScreen "          ", 17, 21
					EmWriteScreen "1", 17, 52
					transmit
					PF3
					EMWaitReady 10, 250
					Exit Do
				End If
			End If
		End If
	
	
	CALL navigate_to_PRISM_screen("SUOL")
	transmit
	Row = Row + 1
	Loop
Loop

MsgBox "All M0935 Worklists have been processed on this caseload! Please review any remaining M0935 Worklists to confirm emancipation!"	


'script_end_procedure("")
script_end_procedure("Success!  " & count & " worklists purged!")

