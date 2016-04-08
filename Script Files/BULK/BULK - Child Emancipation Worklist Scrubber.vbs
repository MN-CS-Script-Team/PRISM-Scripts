'All script files start with the following:
Option Explicit
EMConnect ""

'CONFIRMING PRISM HASN'T TIMED OUT
Dim TimeOutStr
EMWriteScreen "CAST", 21, 18
EMSendKey "<ENTER>"
EMWaitReady 10, 250

'ensures PRISM isn't timed out
EMReadScreen TimeOutStr, 1, 12, 53

If (TimeOutStr = ">") then 
	MsgBox "Please log in first!", vbExclamation
	StopScript
End If

'==================================================================
'		BEGINNING MAIN LOOP FOR M0935 WORKLIST ON USWT		'
'==================================================================

'Declaring Variables used in the loop
Dim M0935Str, M0935_Confirm
Dim Row, Child_Actv, Child_DOB, Child_Age, Child_MCI
Dim SUOD_Type, Child_Row, Child_Col, Emanc_Code

Do 
	EMWriteScreen "USWT", 21, 18
	EMSendKey "<Enter>"
	EMWaitReady 10, 250
	EMWriteScreen "M0935", 20, 30
	EMSendKey "<Enter>"
	EMWaitReady 10, 250

	'CONFIRMING WORKLIST IS M0935
	EMReadScreen M0935Str, 5, 7, 45
	If M0935Str <> "M0935" Then Exit Do		
	EMWriteScreen "D", 7, 4
	EMSendKey "<Enter>"
	EMWaitReady 10, 250

	'CONFIRMING WORKLIST HASN'T ALREADY BEEN REVIEWED
	EMReadScreen M0935_Confirm, 3, 10, 4	
	If M0935_Confirm <> "___" Then Exit Do
	EMWriteScreen "CHDE", 21, 18
	EMSendKey "<ENTER>"
	EMWaitReady 10, 250
	EMWriteScreen "B", 3, 29
	EMSendKey "<ENTER>"
	EMWaitReady 10, 250

	'BEGINNING LOOP TO FIND CHILD
	Row = 8
	Do
		EMReadScreen Child_Actv, 1, Row, 35
		If Child_Actv = " " Then 
			MsgBox "Unable to find child with an 18th birthday within the next 3 months! Please process worklist manually! Script Ended.", VBExclamation
			StopScript
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
	EMWriteScreen "SUOL", 21, 18
	EMSendKey "<ENTER>"
	EMWaitReady 10, 250
	Row = 10
	Do
		'IF UNABLE TO FIND EMANCIPATION LANGUAGE ON ANY ORDER, UPDATING WORKLIST WITH NOTE
		EMReadScreen SUOD_Type, 3, Row, 22
		If SUOD_Type = "   " Then
			EMWriteScreen "USWT", 21, 18
			EMSendKey "<Enter>"
			EMWaitReady 10, 250
			EMWriteScreen "M0935", 20, 30
			EMSendKey "<Enter>"
			EMWaitReady 10, 250
			EMWriteScreen "M", 7, 4
			EMSendKey "<ENTER>"
			EMWaitReady 10, 250
			EMWriteScreen "~.~REVIEWED BY M0935 WORKLIST ON " & Date, 10, 4
			EMWriteScreen "ORDER DOES NOT ADDRESS EMANCIPATION - FURTHER REVIEW NEEDED", 11, 4
			EMWriteScreen "          ", 17, 21
			EmWriteScreen "1", 17, 52
			EMSendKey "<ENTER>"
			EMWaitReady 10, 250
			EMSendKey "<PF3>"
			EMWaitReady 10, 250
			Exit Do
		ElseIf SUOD_Type <> "   " Then 
			EMSetCursor Row, 72
			EMSendKey "<ENTER>"
			EMWaitReady 10, 250
			EMSendKey "<PF11>"
			EMWaitReady 10, 250

			'LOOKING FOR CHILD'S MCI TO CONFIRM EMANCIPATION LANGUAGE
			Child_Row = 1
			Child_Col = 1
			EMSearch Child_MCI, Child_Row, Child_Col

			'IF CHILD'S EMANCIPATION LANGUAGE IS GR, DORD DOC F0300 AND F0302 ARE GENERATED
			If Child_Row > 0 Then
				EMreadScreen Emanc_Code, 2, Child_Row, 66
				If Emanc_Code = "GR" Then
					EMWriteScreen "DORD", 21, 18
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					EMWriteScreen "C", 3, 29
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					EMWriteScreen "A", 3, 29
					EMWriteScreen "F0300", 6, 36
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					Child_Row = 1
					Child_Col = 1
					EMSearch Child_MCI, Child_Row, Child_Col
					EMSetCursor Child_Row, Child_Col
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250

					EMWriteScreen "C", 3, 29
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					EMWriteScreen "A", 3, 29
					EMWriteScreen "F0302", 6, 36
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					Child_Row = 1
					Child_Col = 1
					EMSearch Child_MCI, Child_Row, Child_Col
					EMSetCursor Child_Row, Child_Col
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250

					'PURGING WORKLIST (DOCS TO PRINT OVERNIGHT)
					EMWriteScreen "USWT", 21, 18
					EMSendKey "<Enter>"
					EMWaitReady 10, 250
					EMWriteScreen "M0935", 20, 30
					EMSendKey "<Enter>"
					EMWaitReady 10, 250
					EMWriteScreen "P", 7, 4
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					Exit Do

				'IF CHILD'S EMANCIPATION LANGUAGE IS NOT GR, UPDATING WORKLIST WITH NOTE
				ElseIf Emanc_Code <> "GR" AND Emanc_Code <> "__" Then
					EMWriteScreen "USWT", 21, 18
					EMSendKey "<Enter>"
					EMWaitReady 10, 250
					EMWriteScreen "M0935", 20, 30
					EMSendKey "<Enter>"
					EMWaitReady 10, 250
					EMWriteScreen "M", 7, 4
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					EMWriteScreen "~.~REVIEWED BY M0935 WORKLIST ON " & Date, 10, 4
					EMWriteScreen "ORDER DOES NOT HAVE STANDARD GRADUATION LANGUAGE - FURTHER REVIEW NEEDED", 11, 4
					EMWriteScreen "          ", 17, 21
					EmWriteScreen "1", 17, 52
					EMSendKey "<ENTER>"
					EMWaitReady 10, 250
					EMSendKey "<PF3>"
					EMWaitReady 10, 250
					Exit Do
				End If
			End If
		End If
	
	EMWriteScreen "SUOL", 21, 18
	EMSendKey "<ENTER>"
	EMWaitReady 10, 250
	Row = Row + 1
	Loop
Loop

MsgBox "All M0935 Worklists have been processed on this caseload! Please review any remaining M0935 Worklists to confirm emancipation!"	
