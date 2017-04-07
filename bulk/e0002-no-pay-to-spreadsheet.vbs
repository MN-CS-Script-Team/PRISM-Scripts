'STATS GATHERING ---------------------------
name_of_script = "e0002-no-pay-to-spreadsheet.vbs"
start_time = timer

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
call changelog_update("03/28/2017", "Initial version.", "Kallista Imdieke, Stearns County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'CONNECTING TO BLUEZONE
EMConnect ""

Dim TimeOutStr
EMWriteScreen "CAST", 21, 18
Transmit

EMReadScreen TimeOutStr, 1, 12, 53
If (TimeOutStr = ">") then
	MsgBox "Please log in first!", vbExclamation
	StopScript
End If

'FILTERING FOR E0002 ON USWT
EMWriteScreen "USWT", 21, 18
Transmit
EMWriteScreen "E0002", 20, 30
Transmit

'Opening/Adding into excel worbook 
set objExcel = CreateObject("Excel.Application")
Call excel_open ("H:\Global Applications\Gateway Services\CSU\2017 E0002 No Pays.xlsx", True, True, ObjExcel, objWorkbook)

'BEGINNING WORKLIST LOOP FOR PROCESSING E0002
Dim Worklist_E0002, worklist_date, PRISM_case_number, NCP_name, file_location, last_contact, phone_number

	Do
		'CONFIRMS FIRST WORKLIST ON CAWT LIST IS E0002
		EMReadScreen Worklist_E0002, 5, 7, 45
		IF Worklist_E0002 <> "E0002" Then
			MsgBox ("All E0002 Worklists have been processed on this caseload!")
			Exit Do
		End If

		Row=2
		Do
			If objExcel.Cells(row, 1).Value <> "" THEN row = row + 1
		Loop Until objExcel.Cells(row, 1).Value = ""

		'DISPLAYING E0002 WORKLIST 
		EMWriteScreen "D", 7, 4
		Transmit
		EMReadScreen worklist_date, 10, 17, 21
		objExcel.Cells(row, 1).Value = worklist_date
		EMReadScreen PRISM_case_number, 13, 4, 08
		objExcel.Cells(row, 2).Value = PRISM_case_number
		EMReadScreen NCP_name, 20, 7, 12
		objExcel.Cells(row, 3).Value = NCP_name
		'pulling date of last contact to NCP 
		CALL navigate_to_PRISM_screen("CAAD")
		EMReadScreen file_location, 7, 5, 73
		objExcel.Cells(row, 4).Value = file_location
		write_date DateAdd("m", -3, date), "MM/DD/YYYY", 20, 45
		PF21
		caad_row = 8  
		found = FALSE
		found_once = FALSE
		DO
			EMReadScreen end_of_data, 11, caad_row, 32
			EMReadScreen Type_CAAD, 5, caad_row, 22
			IF Type_CAAD = "T0055" OR Type_CAAD = "T0056" OR Type_CAAD = "T0057" OR Type_CAAD = "T0058" OR Type_CAAD = "T0059" OR Type_CAAD = "M3911" THEN
				EMReadScreen last_contact, 8, caad_row, 12
				found_once = TRUE
			END IF
			IF end_of_data <> "End of Data" THEN
				caad_row = caad_row + 1
			END IF
			IF caad_row = 19 THEN    'Navigate to a new page
				caad_row = 8
				PF8
			END IF
		LOOP UNTIL found = TRUE OR end_of_data = "End of Data"  
		objExcel.Cells(row, 5).Value = last_contact
		CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen phone_number, 12, 14, 14
		objExcel.Cells(row, 6).Value = phone_number
		CALL navigate_to_PRISM_screen("CAWT")
		EMWriteScreen "E0002", 20, 29
		Transmit
		EMWriteScreen "D", 8, 4 
		Transmit
		EMWriteScreen "P", 3, 30 
		Transmit
		Transmit
		PF3

		'AUTOFITTING COLUMNS OF EXCEL SPREADSHEET
		For excel_column = 1 to 6
			objExcel.Columns(excel_column).Autofit()
		Next
		
		'RETURNING TO USWT FOR ADDITIONAL E0002 WORKLISTS
		EMWriteScreen "USWT", 21, 18
		Transmit
		EMWriteScreen "E0002", 20, 30
		Transmit
		EMReadScreen Worklist_E0002, 5, 7, 45
	
		'INCREMENTING EXCEL BY ONE SO IT KEEPS ADDING CASE INFO TO THE SPREADSHEET
		excel_row = excel_row + 1
		Type_CAAD = ""
		last_contact = ""
		

	Loop 

objExcel.ActiveWorkbook.Save 
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

script_end_procedure("")
