'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "iw-caad-cawt.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/22/2017", "Worker Signature should now auto-populate.", "Kelly Hiestand, Wright County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG---------------------------------------------------------------------------
DIM IW_Dialog, PRISM_case_number, Employer_Name, Monthly, Percent, Manual, Manual_Amount, IWType, CAWT, err_msg, ButtonPressed, case_number_is_valid


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
