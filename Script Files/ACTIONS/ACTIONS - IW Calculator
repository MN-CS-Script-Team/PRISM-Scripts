name_of_script = "ACTIONS - IW CALCULATOR.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block 


'this is a function document
DIM beta_agency 'remember to add

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO                                                                          'Declares variables to be good to option explicit users
If beta_agency = "" then                                              'For scriptwriters only
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then                 'For beta agencies and testers
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else                                                                                                                        'For most users
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")                                                               'Creates an object to get a URL
req.open "GET", url, False                                                                                                                                            'Attempts to open the URL
req.send                                                                                                                                                                                                              'Sends request
If req.Status = 200 Then                                                                                                                                                '200 means great success
                Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
                Execute req.responseText                                                                                                                          'Executes the script code
ELSE                                                                                                                                                                                                                       'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
                MsgBox                "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
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
'this is where the copy and paste from functions library ended



'DIALOGS---------------------------------------------------------------------------
DIM IW_CALC_Dialog, PRISM_case_number, Current_Support, Percent, Manual, Other_Amount, err_msg, ButtonPressed, case_number_is_valid, MoTotal, Month_NonAccrual, Month_Accrual

BeginDialog IW_CALC_Dialog, 0, 0, 176, 110, "IW CALC Dialog"
  EditBox 60, 5, 95, 15, PRISM_case_number
  EditBox 100, 25, 55, 15, Current_Support
  CheckBox 10, 65, 45, 10, "20 Percent", Percent
  ButtonGroup ButtonPressed
    OkButton 60, 85, 50, 15
    CancelButton 120, 85, 50, 15
  Text 5, 10, 50, 10, "Case Number"
  Text 5, 30, 90, 10, "Current Monthly Obligation "
  Text 5, 50, 80, 10, "Arrears Collection Rate"
EndDialog



'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

CALL navigate_to_PRISM_screen ("CAFS")

'variable name in edit box that i want autofilled
EMReadScreen PRISM_case_number, 13, 4, 8

'getting amounts to autofill
EMReadScreen Month_Accrual, 8, 9, 31
EMReadScreen Month_NonAccrual, 8, 10, 31
Month_Accrual = Trim(Month_Accrual)
Month_NonAccrual = Trim(Month_NonAccrual)

'Converting accrual amts to number from string and calculating total monthly amount
Month_Accrual = CDbl(Month_Accrual)
Month_NonAccrual = CDbl(Month_NonAccrual)


Current_Support = Month_Accrual + Month_NonAccrual 
Current_Support = Trim(Current_Support)
Current_Support = FormatNumber(Current_Support)

'adding a loop
Do
	err_msg = ""
	Dialog IW_CALC_Dialog				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed."
		IF Current_Support = "" THEN err_msg = err_msg & vbNewline & "Current Support must be completed"
		'IF CP = 0 AND NCP = 0 THEN err_msg = vbNewline & "Either CP or NCP must be selected."

		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & "Please resolve for the script to continue."
		END IF
LOOP UNTIL err_msg = ""


Current_Support = CDbl(Current_Support)

IF Percent = checked THEN MoTotal = Current_Support * 1.2
IF Percent = 0 THEN MoTotal = Current_Support

'Calculating pay period amounts
Dim WeekPay, BiWeekPay, SemiMoPay

WeekPay = MoTotal /4.333
WeekPay = FormatNumber(WeekPay, 2)


BiWeekPay = MoTotal /2.167
BiWeekPay = FormatNumber(BiWeekPay, 2)

SemiMoPay = MoTotal/2
SemiMoPay = FormatNumber(SemiMoPay, 2)

MoTotal = FormatNumber(Mototal)

'takes you to palc so you can see the amount that is being received on the case
CALL navigate_to_PRISM_screen ("PALC")
 

'msgbox needed to show calculations, weekly, biweekly, semi monthly, and monthly with 20%
IF Percent = checked THEN 
	MsgBox ("Monthly: $" & MoTotal & VbNewline & VbNewline & _
		"Weekly: $" & WeekPay & VbNewline & VbNewline & _
		"Bi-Weekly: $" & BiWeekPay & VbNewline & VbNewline & _
		"Semi-Monthly: $" & SemiMoPay & VbNewline & VbNewline & _
		"20% of current support:  $" & Current_Support * .2)

END IF

'without 20%
IF Percent = 0 THEN 
	MsgBox ("Monthly: $" & MoTotal & VbNewline & VbNewline & _
		"Weekly: $" & WeekPay & VbNewline & VbNewline & _
		"Bi-Weekly: $" & BiWeekPay & VbNewline & VbNewline & _
		"Semi-Monthly: $" & SemiMoPay) 


END IF


script_end_procedure("")
