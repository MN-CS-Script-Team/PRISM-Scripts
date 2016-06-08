'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "BULK - FAILURE POF - RSDI DFAS.vbs" 
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
Call navigate_to_Prism_screen("USWT")

' >>>>> SELECTING THE SPECIFIC WORKLIST TYPE <<<<<
EMWriteScreen "E0014", 20, 30
transmit

USWT_row = 7
COUNT = 0
SCROLL = 0
' >>>>> STARTING THE DO LOOP. THE SCRIPT NEEDS TO HANDLE THESE CASES ONE AT A TIME <<<<<
DO
	EMReadScreen USWT_type, 5, USWT_row, 45
	IF USWT_type = "E0014" THEN
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		EMWriteScreen "s", USWT_row, 4
		transmit
		'Selecting the worklist brings the user to NCP's PAPL screen
		
		purge = false 'Reset the purge variable

		' >>>>> MAKING SURE THAT THERE IS INFORMATION ON PAPL <<<<
		EMReadScreen end_of_data, 11, USWT_row, 32
		IF end_of_data <> "End of Data" THEN

			' >>>>> READING THE MOST RECENT PAY DATE AND CONVERTING IT TO A USABLE DATE <<<<<
			EMReadScreen PAPL_most_recent_pay_date, 6, 7, 7
			Call date_converter_PALC_PAPL(PAPL_most_recent_pay_date)
			pmt_year = Right(PAPL_most_recent_pay_date, 2) 'string variables added to track the payment month and 2-digit year.
			pmt_month = Left(PAPL_most_recent_pay_date, 2)	
			
						
			' >>>> CHECKING THAT THE DATE IN THE PAYMENT ID IS FROM THE CURRENT MONTH MINUS 1 <<<<<
			current_month_minus1 = DateAdd("m", -1, date) 'variable for the current date minus one - this returns a date format
			c_month = datepart("m", current_month_minus1)
			IF len(c_month) = 1 THEN c_month = "0" & c_month
			
			
			c_year = Right(CStr(current_month_minus1), 2) 'string variables added to track the current month minus 1 month and year. 
			'c_month = Left(CStr(current_month_minus1), 2)
			
			IF pmt_year >= c_year THEN
				If  pmt_month >= c_month THEN  
 				' >>>>> IF THE PAYMENT IS FROM LAST MONTH OR CURRENT MONTH, THE SCRIPT GRABS THE EMPLOYER/SOURCE ID <<<<<
				'We want this to occur if the payment occurred last month or in the current month.				
					PF11
					EMReadScreen PAPL_name, 30, 7, 38
					' >>>>> LISTING OUT THE CONDITIONS THAT CAN BE PURGED AUTOMATICALLY <<<<<
					IF InStr(PAPL_name, "DFAS") <> 0 OR _
					   InStr(PAPL_name, "U S SOCIAL") <> 0 OR _ 
					   InStr(PAPL_name, "U S DEPT OF TREASURY") <> 0 THEN 
						purge = True
					 	COUNT = COUNT + 1
					   	Msgbox USWT_case_number & " worklist selected for purge!"
					Else
						purge = false
					End If
				End If
			END IF
		End If

		
		Call navigate_to_PRISM_screen ("CAWT")
		EMWriteScreen "E0014", 20, 29
		EMWriteScreen USWT_case_number, 20, 8
		transmit

		' >>>>> IF THE WORKLIST ITEM IS ELIGIBLE TO BE PURGED, THE SCRIPT PURGES...
		IF purge = True THEN 
			CAWT_row = 8
			DO 
				EMReadScreen CAWD_type, 5, cawt_row, 8
				If cawd_type = "E0014" then	
					EMWriteScreen "P", caWT_row, 4
					transmit
					transmit
					PF3
				End if
				cawt_row = cawt_row + 1
			LOOP until cawd_type <> "E0014"
		END IF
		'  ...  IF THE WORKLIST ITME IS NOT ELIGIBLE TO BE PURGED, THE SCRIPT INCREASES USWT_ROW + 1 <<<<<
			Call navigate_to_PRISM_screen ("USWT")

			EMWriteScreen "E0014", 20, 30
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
		
	End If
LOOP UNTIL USWT_type <> "E0014"

script_end_procedure("Success!  " & Count & " worklists purged!")
