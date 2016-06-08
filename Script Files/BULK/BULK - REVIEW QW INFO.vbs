'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "BULK - REVIEW QW INFO.vbs" 
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

FUNCTION quarterly_wage(participant)
	'Setting variables
	IF participant = "CP" then
		employer_screen = "CPID"
		worklist = "L2501"
	END IF
	IF participant = "NCP" then
		employer_screen = "NCID"
		worklist = "L2500"
	END IF 
	

''	CALL navigate_to_Prism_screen("REGL")
'	transmit

'>>>>> GOING TO USWT <<<<<
CALL navigate_to_Prism_screen("USWT")

' >>>>> SELECTING THE SPECIFIC WORKLIST TYPE <<<<<
EMWriteScreen worklist, 20, 30
transmit

USWT_row = 7
count = 0
SCROLL = 0

' >>>>> STARTING THE DO LOOP. THE SCRIPT NEEDS TO HANDLE THESE CASES ONE AT A TIME <<<<<
'Creating a placeholder string to check that the case we are working on has not already been worked on.
'This will prevent the script from getting stuck on cases that are not purged.\
placeholder_case_number_string = ""
DO
	EMReadScreen USWT_type, 5, USWT_row, 45
	IF USWT_type = worklist THEN
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		'If the script does not find this case number in the placeholder string, we will build on that string,
		'and we will go ahead with the logic to check the quarterly wage match on this case.

		IF InStr(placeholder_case_number_string, USWT_case_number) = 0 THEN 
			placeholder_case_number_string = placeholder_case_number_string & "~~~" & USWT_case_number
			EMWriteScreen "s", USWT_row, 4
			transmit
			'Selecting the worklist brings the user to the quarterly wage browse screen
			'Need to go through the page of to locate the unreviewed results
			QW_row = 9
			placeholder_qw_string = ""


			DO 
				
				EMReadScreen end_of_data_check, 11, QW_row, 32
				IF end_of_data_check = "End of Data" THEN EXIT DO	
				EMReadScreen rev_check, 1, QW_row, 75
		
			'When an unreviewed result is found, need to display it.
				IF rev_check <> "Y" THEN
					EMReadScreen qw_string, 65, QW_row, 8
					IF InStr(placeholder_qw_string, qw_string) = 0 THEN
						placeholder_qw_string = placeholder_qw_string & "~~~" & qw_string
					'msgbox placeholder_qw_string

					EMWriteScreen "D", QW_row, 4
					transmit
					'Then need to hit F6 to update
					PF6
					'If employer already exists on employer screen, mark the case reviewed.
					
					EMReadScreen bottom_line_message, 70, 24, 3
					bottom_line_message = trim(bottom_line_message)
					IF bottom_line_message <> "" THEN 
						IF InStr(bottom_line_message, "already exists") <> 0 THEN 
							PF3
							EMWriteScreen "M", 3, 29  	'Modify the page
							EMWriteScreen "Y", 16, 64     'Mark reviewed
							count = count + 1
							transmit
							PF3   'return to the qw screen
						'	msgbox USWT_case_number & QW_row & "- This one will be worked by the script."
						ELSEIF InStr(bottom_line_message, "pf6 to select") <> 0 THEN 
							PF3
							PF3
						ELSEIF InStr(bottom_line_message, "Fein is required") <> 0 THEN
							PF3	
				'If the employer is new, prompt the user if they want to add it.  If they want to add it, mark the case reviewed
				'	continue = Msgbox("Attempt to add new employer to this participant's employer screen"_
				'				"and mark this wage match reviewed?", 4, "Add Employer?")	
				'	IF continue = 6 then 'User selected to add the employer
				'	
				'	END IF	
				'	IF continue = 7 then 'User selected not to add the employer
				'		PF3
				'	END IF
				'		msgbox USWT_case_number &  " " & QW_row & "- This needs to be reviewed by the user, new employer."
						ELSE 'Some other message is displayed
							PF3
							PF3
						END IF
					ELSEIF bottom_line_message = "" THEN 
							PF3
							PF3
					'		msgbox USWT_case_number & " "& QW_row & "- This needs to be reviewed by the user."
					'		
					END IF
				END IF
			'If the employer does not meet the above, leave the case un-reviewed.
				END IF			
				QW_row = QW_row + 1
				IF QW_row = 19 THEN      	'Pagination
					PF8
					QW_row = 9
				END IF
			
			LOOP UNTIL end_of_data_check = "End of Data"
	
			'Advances to the next case 		
			CALL navigate_to_PRISM_screen ("USWT")
			EMWriteScreen worklist, 20, 30
			transmit
		END IF	
			
		USWT_row = USWT_row + 1
		EMReadScreen end_of_data, 11, uswt_row, 32
		IF end_of_data = "End of Data" THEN EXIT DO
		
		IF USWT_row = 19 THEN      	'Pagination
			PF8
			USWT_row = 7
		END IF
		
	END IF
LOOP UNTIL USWT_type <> worklist


MsgBox count & " quarterly wage match records for " & participant & " have been reviewed by the script!"


END FUNCTION

' >>>>> THE SCRIPT <<<<<
EMConnect ""

CALL quarterly_wage("NCP")
CALL quarterly_wage("CP")
CALL quarterly_wage("NCP")
CALL quarterly_wage("CP")
script_end_procedure("Success! The script is now ending!")
