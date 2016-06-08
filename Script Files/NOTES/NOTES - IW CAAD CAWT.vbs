'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IW CAAD AND CAWT.vbs"
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
'this is where the copy and paste from functions library ended



'DIALOG---------------------------------------------------------------------------
DIM IW_Dialog, PRISM_case_number, Employer_Name, Monthly, Percent, Manual, Manual_Amount, IWType, CAWT, worker_signature, err_msg, ButtonPressed, case_number_is_valid


BeginDialog IW_Dialog, 0, 0, 201, 180, "IW CAAD CAWT CALC Dialog"
  EditBox 60, 5, 110, 15, PRISM_case_number
  EditBox 65, 30, 105, 15, Employer_Name
  EditBox 115, 55, 55, 15, Monthly
  DropListBox 50, 75, 60, 45, "Select one..."+chr(9)+"New"+chr(9)+"Amended"+chr(9)+"EIWO", IWType
  CheckBox 5, 100, 135, 10, "Manual IW sent.  Arrears collection is", Manual_Amount
  EditBox 145, 95, 50, 15, Manual
  CheckBox 5, 120, 195, 10, "Check here to create a follow up CAWT note 30 days out.", CAWT
  EditBox 80, 135, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 85, 160, 50, 15
    CancelButton 145, 160, 50, 15
  Text 5, 10, 50, 10, "Case Number"
  Text 5, 35, 55, 10, "Employer Name"
  Text 5, 60, 105, 10, "Monthly Collection on IW Notice "
  Text 5, 80, 40, 10, "Type of IW"
  Text 5, 140, 70, 10, "Worker Signature"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'to pull up my prism 
EMFocus

'checks to make sure we are in PRISM
CALL check_for_PRISM(True)

'taking me to cast so i can read the case number to put in dialog box
CALL navigate_to_PRISM_screen ("CAST")

'it is reading the case number and putting in dialog box
EMReadScreen PRISM_case_number, 13, 4, 8 

'taking me to cast so i can read the employer to put in dialog box
CALL navigate_to_PRISM_screen ("NCID")

EMWriteScreen "B", 3, 29
Transmit

'it is reading the case number and putting in dialog box
EMReadScreen Employer_Name, 20, 8, 51 

'Calculating pay period amounts to put in cawt and caad
Dim total_arrears, Month_Accrual, Month_NonAccrual

'***********want monthly to be mo accrual plus month non accural and auto fill in dialog

EMWriteScreen "CAFS", 21, 18
EMSendKey "<Enter>"
EMWaitReady 10, 250
EMWaitForText "** Case Balances **", 8, 31, 30
EMReadScreen total_arrears, 10, 12, 68
total_arrears = Trim(total_arrears)

EMReadScreen Month_Accrual, 8, 9, 31
EMReadScreen Month_NonAccrual, 8, 10, 31
Month_Accrual = Trim(Month_Accrual)
Month_NonAccrual = Trim(Month_NonAccrual)

'making sure script read numbers as number not strings
Monthly = Monthly * 1
Month_Accrual = Month_Accrual * 1
Month_NonAccrual = Month_NonAccrual * 1
total_arrears = total_arrears * 1

'calculating monthly collection to put in dialog and caad and cawt
IF total_arrears = 0 THEN Monthly = Month_Accrual + Month_NonAccrual  
IF total_arrears >= Month_Accrual AND Month_NonAccrual = 0  THEN Monthly = (Month_Accrual + Month_NonAccrual) * 1.2
IF total_arrears >= Month_Accrual AND Month_NonAccrual > 0  THEN Monthly = (Month_Accrual + Month_NonAccrual)
IF total_arrears > Month_NonAccrual AND Month_Accrual = 0 THEN Monthly = Month_NonAccrual * 1.2
IF total_arrears < Month_Accrual AND total_arrears <> 0 AND Month_NonAccrual = 0 THEN Monthly = Month_Accrual
IF total_arrears < Month_Accrual AND total_arrears <> 0 AND Month_NonAccrual > 0  THEN Monthly = (Month_Accrual + Month_NonAccrual) 

Monthly = trim(Monthly)

'formating to currency with $
Monthly = FormatCurrency(Monthly)
Month_Accrual = FormatCurrency(Month_Accrual)
Month_NonAccrual = FormatCurrency(Month_NonAccrual)
total_arrears = FormatCurrency(total_arrears)

'***************************************

'THE LOOP----------------------------------------
'adding a loop
Do
	err_msg = ""	
	Dialog IW_Dialog	'shows name of dialog		
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed"
		IF Monthly = "" THEN err_msg = err_msg & vbNewline & "Total monthly Collection on IW Notice must be completed."
		IF Employer_Name = "" THEN err_msg = err_msg & vbNewline & "Employer Name must be completed."
		IF IWType = "Select one..." THEN err_msg = err_msg & vbNewline & "IW Type must be completed.  "
		IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "Please sign your CAAD Note."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'----------------------------------------------------
'Calculating pay period amounts to put in cawt and caad
Dim WeekPay, BiWeekPay, SemiMoPay


WeekPay = Monthly /4.333
WeekPay = FormatNumber(WeekPay, 2)
BiWeekPay = Monthly /2.167
BiWeekPay = FormatNumber(BiWeekPay, 2)
SemiMoPay = Monthly/2
SemiMoPay = FormatNumber(SemiMoPay, 2)


'brings me to caad and creates a FREE note
CALL navigate_to_PRISM_screen ("CAAD")
PF5
EMWriteScreen "A", 3, 29
EMWriteScreen "free", 4, 54
EMSetCursor 16, 4

'this will add information to the caad note
		CALL write_variable_in_CAAD ("*" & IWType & " IW sent to " & Employer_Name  & Monthly  & " per month")
IF Manual_Amount = checked THEN CALL write_variable_in_CAAD ("*Manual IW sent. Arrears collection is $" & Manual)
CALL write_variable_in_CAAD ("weekly: $" & WeekPay & "  biweekly: $" & BiWeekPay & "  semimonthly: $"& SemiMoPay)
CALL write_variable_in_CAAD(worker_signature)
transmit
PF3

'creating CAWT note 30 days out
IF CAWT = checked THEN
CALL navigate_to_PRISM_screen ("CAWT")
PF5
EMWriteScreen "free", 4, 37
EMSetCursor 10, 4
	CALL write_variable_in_CAAD ("Did " & IWType & " IW start from "  &  Employer_Name  &  Monthly  &  " per month"  &  " yet?")
CALL write_variable_in_CAAD ("weekly: $" & WeekPay & "  biweekly: $" & BiWeekPay & "  semimonthly: $"& SemiMoPay)
EMWriteScreen "30", 17, 52 
transmit
PF3
End IF

script_end_procedure("")


