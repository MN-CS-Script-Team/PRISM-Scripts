'Gathering stats-------------------------------------------------------------------------------------
name_of_script = "BULK - NCP COMPANION CASE FINDER.vbs"
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

' >>>>> BUILDING THE DIALOG.
'       THIS DIALOG IS USED TO GATHER THE POSITION NUMBERS. <<<<<
BeginDialog get_cali_location_dlg, 0, 0, 221, 95, "Enter Position Number"
  EditBox 10, 45, 205, 15, position_number
  ButtonGroup ButtonPressed
    OkButton 115, 75, 50, 15
    CancelButton 165, 75, 50, 15
  Text 10, 10, 200, 20, "Please enter 11-digit Position Number. You can enter multiple if you separate them with a comma."
EndDialog


' >>>>> THE SCRIPT <<<<<
EMConnect ""

' >>>>> BUILDING IN SAFEGUARD TO MAKE SURE THE WORKER IS ENTERING AN 8-DIGIT WORKER NUMBER OR AN 11-DIGIT POSITION NUMBER. <<<<<
DO
	err_msg = ""
	' >>>>> CALLING THE DIALOG <<<<<
	DIALOG get_cali_location_dlg
		IF ButtonPressed = 0 THEN stopscript
		CALL check_for_PRISM(false)
		position_number = replace(position_number, " ", "")
		position_number = split(position_number, ",")
		
		' >>>>> BUILDING ERROR MESSAGE FOR ALL POSITION NUMBERS THAT ARE NOT 11 DIGITS LONG. <<<<<
		FOR EACH worker_position IN position_number
			IF worker_position <> "" THEN 
				IF len(worker_position) <> 11 AND len(worker_position) <> 8 THEN err_msg = err_msg & vbCr & "* Position: " & worker_position & " is not a valid 8-digit 11-digit number."
			END IF
		NEXT
		
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

' >>>>> BUILDING A NEW EXCEL FILE FOR EACH CALI. <<<<<
FOR EACH worker_position IN position_number
	IF worker_position <> "" THEN 
		' >>>>> BUILDING THE CONFIRMATION MESSAGE FOR THIS worker_position.
		'       THE SCRIPT CAN DIFFERENTIATE BETWEEN 8-DIGIT WORKER NUMBERS AND 11-DIGIT POSITION NUMBERS. <<<<<
		IF len(worker_position) = 8 THEN 
			'If the length of the worker number is 8 then the script goes to CALI to gather the 11-digit worker position number.
			CALL navigate_to_PRISM_screen("CALI")
			CALL check_for_PRISM(False)
			EMWriteScreen left(worker_position, 3), 20, 18
			EMWriteScreen "001", 20, 30
			transmit
			
			EMSetCursor 20, 49
			EMSendKey "X"
		
			PF1
			
			CALI_row = 13
			DO
				EMReadScreen worker_id, 8, CALI_row, 39
				EMReadScreen end_of_data, 11, CALI_row, 39
				IF end_of_data = "End of Data" THEN 
					worker_position_name = "WORKER NOT FOUND"
					EXIT DO
				END IF
				
				IF UCASE(worker_id) = UCASE(worker_position) THEN 
					EMReadScreen worker_position_name, 30, CALI_row, 49
					worker_position_name = trim(worker_position_name)
					EMReadScreen CALI_position, 20, CALI_row, 18
					CALI_position = replace(CALI_position, " ", "")
					worker_position_position_number = CALI_position
					PF3
					EXIT DO
				ELSE
					CALI_row = CALI_row + 1
					IF CALI_row = 19 THEN 
						PF8
						CALI_row = 13
					END IF
				END IF
			LOOP
		ELSEIF len(worker_position) = 11 THEN
			CALL navigate_to_PRISM_screen("CALI")
			EMSetCursor 20, 18
			EMSendKey worker_position
			transmit
			
			EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
			error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
			IF error_message_on_bottom_of_screen = "" THEN 
				CALL find_variable("Name: ", worker_position_name, 30)
				worker_position_name = trim(worker_position_name)
				worker_position_position_number = worker_position
			ELSEIF error_message_on_bottom_of_screen <> "" THEN 
				worker_position_name = "WORKER NOT FOUND"
			END IF
		ELSEIF len(worker_position) <> 8 AND len(worker_position) <> 11 THEN 
			worker_position_name = "WORKER NOT FOUND"
		END IF
		
		IF worker_position_name = "WORKER NOT FOUND" THEN 
			MsgBox "* Worker position " & worker_position & " cannot be found. Skipping this ID."			
		ELSE
			confirmation_message = MsgBox("*** PLEASE CONFIRM ***" & vbCr & vbCr & "Please confirm that you are building a list for: " & worker_position_name & " at position " & worker_position & "." & vbCr & vbCr & "Press YES to proceed. Press NO to skip this position. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirmation_message = vbCancel THEN stopscript
			
			
			IF confirmation_message = vbYes THEN 
				CALL check_for_PRISM(false)
				IF len(worker_position) <> 11 THEN worker_position = CALI_position	
				' >>>>> CREATING THE EXCEL FILE <<<<<
				' >>>>> MAKING THE EXCEL FILE NOT VISIBLE UNTIL THE SCRIPT IS DONE. THIS PREVENTS WORKERS FROM CLICKING AND INTERRUPTING THE FLOW. <<<<<
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = False
				Set objWorkbook = objExcel.Workbooks.Add()
				objExcel.DisplayAlerts = True
				objExcel.Caption = "Companion Cases for " & worker_position
				
				' >>>>> EXCEL COLUMN HEADERS
				objExcel.Cells(1, 1).Value = "CASE NUMBER"
				objExcel.Cells(1, 1).Font.Bold = TRUE
				objExcel.Cells(1, 2).Value = "NCP MCI"
				objExcel.Cells(1, 2).Font.Bold = TRUE
				objExcel.Cells(1, 3).Value = "NCP NAME"
				objExcel.Cells(1, 3).Font.Bold = TRUE
				objExcel.Cells(1, 4).Value = "COMPANION CASES (with Worker ID)"
				objExcel.Cells(1, 4).Font.Bold = TRUE
				
				' >>>>> AUTOFITTING THE COLUMNS FOR MAKING PRETTY. <<<<<
				FOR i = 1 TO 4
					IF i = 2 THEN 
						objExcel.Columns(i).ColumnWidth = 10.71
					ELSE
						objExcel.Columns(i).AutoFit()
					END IF
				NEXT
				
				excel_row = 2
				
				' >>>>> RESETING PRISM <<<<<
				CALL navigate_to_PRISM_screen("REGL")
				transmit
				
				' >>>>> NAVIGATING TO CALI TO START GRABBING CASE NUMBERS AND NCP MCIs <<<<<
				CALL navigate_to_PRISM_screen("CALI")
				EMSetCursor 20, 18
				EMSendKey worker_position
				transmit
				
				' >>>>> MAKING SURE THE SCRIPT GETS TO THE TOP OF THE CALI. <<<<<
				' >>>>> THIS IS NECESSARY IF THE WORKER STARTS FROM THEIR OWN CALI. <<<<<
				DO
					PF7
					EMReadScreen top_of_scroll, 13, 24, 2
				LOOP UNTIL top_of_scroll = "Top of scroll"
				
				' >>>>> GETTING TO PAGE 2 TO GRAB NCP INFORMATION. <<<<<
				DO
					EMReadScreen ncp_name, 8, 6, 35
					IF ncp_name <> "NCP Name" THEN PF11
				LOOP UNTIL ncp_name = "NCP Name"
				
				CALI_row = 8
				DO
					' >>>>> READING MCI <<<<<
					EMReadScreen NCP_MCI, 10, CALI_row, 22
					NCP_MCI = CStr(NCP_MCI)
					' >>>>> READING CASE NUMBER <<<<<
					EMReadScreen PRISM_case_number, 14, CALI_row, 7
					PRISM_case_number = replace(PRISM_case_number, " ", "")
					PRISM_case_number = CStr(PRISM_case_number)
					' >>>>> GRABBING NCP NAME <<<<<
					EMReadScreen NCP_name, 30, CALI_row, 33
					NCP_name = trim(NCP_name)
					EMReadScreen end_of_data, 11, CALI_row, 32
					' >>>>> CHECKING THAT THE SCRIPT IS NOT AT THE END OF CALI... <<<<<
					IF end_of_data <> "End of Data" THEN 
						' >>>>> CHECKING THAT THE NCP IS KNOWN... <<<<<
						IF NCP_MCI <> "          " THEN 
							' >>>>> CHECKING THAT THE USER CAN GRAB INFORMATION ABOUT THIS CASE... <<<<<
							IF InStr(NCP_name, "Access Denied") = 0 THEN 
								' >>>>> MODIFYING THE CASE NUMBER AND MCI FOR DISPLAY IN EXCEL <<<<<
								' >>>>> THEN ADDING THE CASE NUMBER, MCI, AND NCP NAME TO EXCEL <<<<<
								objExcel.Cells(excel_row, 1).Value = left(PRISM_case_number, 10) & "-" & right(PRISM_case_number, 2)
								DO
									IF len(NCP_MCI) <> 10 THEN NCP_MCI = "0" & NCP_MCI
								LOOP UNTIL len(NCP_MCI) = 10
								objExcel.Cells(excel_row, 2).Value = CStr(NCP_MCI)
								objExcel.Cells(excel_row, 3).Value = NCP_name			
								excel_row = excel_row + 1
							END IF
						END IF
					ELSEIF end_of_data = "End of Data" THEN 
						' >>>>> IF THE SCRIPT FINDS "END OF DATA" ON CALI_ROW, IT EXITS THE DO/LOOP <<<<<
						EXIT DO
					END IF
					CALI_row = CALI_row + 1
					' >>>>> IF THE SCRIPT GETS TO THE BOTTOM OF CALI, IT WILL TURN THE PAGE AND RESET CALI_ROW TO THE TOP OF THE PAGE <<<<<
					IF CALI_row = 19 THEN 
						PF8
						CALI_row = 8
					END IF
				LOOP UNTIL end_of_data = "End of Data"
				
				objExcel.Columns(3).AutoFit()
				
				' >>>>> RESETTING EXCEL_ROW TO GET THE SCRIPT TO START PULLING NCP MCIs FROM THE TOP OF THE LIST <<<<<
				excel_row = 2
				' >>>>> NAVIGATING TO NCCB <<<<<
				CALL navigate_to_PRISM_screen("NCCB")
				DO
					' >>>>> THE VARIABLE lead_case_number IS BEING USED FOR COMPARISON LATER IN THE SCRIPT.
					'       THE SCRIPT USES lead_case_number TO PREVENT ADDING THE CURRENT CASE NUMBER TO THE LIST OF COMPANION CASES.
					'       THE IDEA BEING THAT IF CASE NUMBERS OTHER THAN THE ONE FOUND ON CALI ARE OPEN IN THE AGENCY AND THE NCP IS THE NCP ON THOSE CASES. <<<<<
					lead_case_number = objExcel.Cells(excel_row, 1).Value
					NCP_MCI = objExcel.Cells(excel_row, 2).Value
					' >>>>> CONVERTING THE NCP MCI TO 10 DIGITS LONG SO IT CAN BE ENTERED ON NCCB. <<<<<
					DO
						IF len(NCP_MCI) <> 10 THEN NCP_MCI = "0" & NCP_MCI 
					LOOP UNTIL len(NCP_MCI) = 10
					' >>>>> ENTER THE NCP MCI AT 20,6 AND TRANSMITTING <<<<<
					CALL write_value_and_transmit(NCP_MCI, 20, 6)
					
					' >>>>> STARTING THE SEARCH FOR COMPANION CASES AT NCCB ROW 7. <<<<<
					NCCB_row = 7
					' >>>>> RESETING THE VALUE FOR THE STRING OF COMPANION CASES. <<<<<
					companion_cases = ""
					DO
						EMReadScreen NCCB_status, 3, NCCB_row, 68
						EMReadScreen NCCB_role, 3, NCCB_row, 8
						EMReadScreen NCCB_worker, 8, NCCB_row, 73
						EMReadScreen end_of_data, 11, NCCB_row, 32
						EMReadScreen NCCB_case_number, 13, NCCB_row, 15
						NCCB_case_number = replace(NCCB_case_number, " ", "-")
						NCCB_case_number = CStr(NCCB_case_number)
						' >>>>> CHECKING THAT THE SCRIPT IS NOT AT THE END OF NCCB... <<<<<
						IF end_of_data <> "End of Data" THEN 
							' >>>>> CHECKING THAT THE INDIVIDUAL IS THE NCP ON THAT CASE...
							'       CHECKING THAT THE INDIVIDUAL IS OPEN ON THE SELECTED CASE...	
							'       CHECKING THAT THE CASE IS ACTIVE IN THE MAINTAINING AGENCY AS THE PRIMARY CASE...
							'		CHECKING THAT THE SELECTED IS UNIQUE FROM THE PRIMARY CASE...
							'       AND ADDING THIS CASE AND THE MAINTAINING WORKER TO THE LIST OF COMPANION CASES. <<<<<			
							IF NCCB_role = "NCP" AND NCCB_status = "OPN" AND left(NCCB_worker, 3) = left(worker_position, 3) AND lead_case_number <> NCCB_case_number THEN companion_cases = companion_cases & NCCB_case_number & " (" & NCCB_worker & "), "
						ELSEIF end_of_data = "End of Data" THEN 
							' >>>>> IF "End of Data" IS FOUND, THE SCRIPT EXITS THE DO/LOOP <<<<<
							EXIT DO
						END IF 	
						NCCB_row = NCCB_row + 1
						IF NCCB_row > 19 THEN 
							PF8
							NCCB_row = 7
						END IF
					LOOP UNTIL end_of_data = "End of Data"
					objExcel.Cells(excel_row, 4).Value = companion_cases
									
					' >>>>> DELETING THE EXCEL ROW IF NO COMPANION CASES ARE FOUND <<<<<
					IF objExcel.Cells(excel_row, 4).Value = "" THEN
						SET objRange = objExcel.Cells(excel_row, 1).EntireRow
						objRange.Delete
						excel_row = excel_row - 1
					END IF	
								
					excel_row = excel_row + 1
				LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""
				
				' >>>>> AUTO-FITTING THE COLUMN WIDTH WHEN FINISHED <<<<<
				objExcel.Columns(4).AutoFit()
				' >>>>> MAKING THE EXCEL FILE VISIBLE <<<<<
				objExcel.Visible = True
				
				' >>>>> REMOVING DUPLICATE NCP MCIs <<<<<
				excel_row = 2
				MCI_array = ""
				DO
					NCP_MCI = objExcel.Cells(excel_row, 2).Value
					IF InStr(MCI_array, NCP_MCI) = 0 THEN 
						MCI_array = MCI_array & NCP_MCI & "~"
					ELSEIF InStr(MCI_array, NCP_MCI) <> 0 THEN 
						SET objRange = objExcel.Cells(excel_row, 1).EntireRow
						objRange.Delete
						excel_row = excel_row - 1
					END IF
					excel_row = excel_row + 1
				LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""
			END IF
		END IF
	END IF
NEXT

script_end_procedure("Success!!")
