

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "interview-information-sheet.vbs"
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
call changelog_update("03/31/2017", "Bug fix for FormatCurrency error message.", "Wendy LeVesseur, Anoka County")
call changelog_update("03/09/2017", "Replace username with worker signature and fix other bugs.", "Wendy LeVesseur, Anoka County")
call changelog_update("12/08/2016", "Initial version.", "Wendy LeVesseur, Anoka County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog


BeginDialog Interview_Info_dialog, 0, 0, 206, 85, "Interview Information Sheet"
  ButtonGroup ButtonPressed
    OkButton 30, 60, 50, 15
    CancelButton 95, 60, 50, 15
  Text 5, 10, 105, 20, "Which participant do you want to prepare the sheet for?"
  DropListBox 115, 10, 80, 20, "NCP"+chr(9)+"CP", participant
  EditBox 115, 35, 80, 15, worker_signature
  Text 45, 40, 60, 10, "Worker Signature:"
EndDialog


'------Start of Class definitions--------------------------------------------------------------------------------
'>>>>> CLASSES!!!!!!!!!!!!!!!!!!!!! <<<<<
' This CLASS contains properties used to populate documents
' These properties should not be used for other applications in scripts.
' Every time you call the property, the script will use the class definition to efficiently obtain the requested information.
CLASS doc_info
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' >>>>> CP INFORMATION <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' CP name (last, first middle initial, suffix (if any))
	PUBLIC PROPERTY GET cp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_name, 50, 5, 25
		cp_name = trim(cp_name)
	END PROPERTY
	
	' CP first name
	PUBLIC PROPERTY GET cp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_first_name, 12, 8, 34
		cp_first_name = trim(replace(cp_first_name, "_", ""))
	END PROPERTY

	' CP last name
	PUBLIC PROPERTY GET cp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_last_name, 17, 8, 8
		cp_last_name = trim(replace(cp_last_name, "_", ""))
	END PROPERTY	
	
	' CP middle name
	PUBLIC PROPERTY GET cp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_middle_name, 12, 8, 56
		cp_middle_name = trim(replace(cp_middle_name, "_", ""))
	END PROPERTY
	
	' CP middle initial
	PUBLIC PROPERTY GET cp_middle_initial
		cp_middle_initial = left(cp_middle_name, 1)
	END PROPERTY
	
	' CP suffix
	PUBLIC PROPERTY GET cp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")	
		EMReadScreen cp_suffix, 3, 8, 74
		cp_suffix = trim(replace(cp_suffix, "_", ""))
	END PROPERTY
	
	' CP date of birth
	PUBLIC PROPERTY GET cp_dob
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_dob, 8, 6, 24		
	END PROPERTY

	' CP social security number
	PUBLIC PROPERTY GET cp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_ssn, 11, 6, 7
	END PROPERTY
	
	' CP MCI
	PUBLIC PROPERTY GET cp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_mci, 10, 5, 7
	END PROPERTY	
	
	' CP address
	PUBLIC PROPERTY GET cp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadscreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_addr1, 30, 15, 11
			EMReadScreen cp_addr2, 30, 16, 11
			cp_addr = replace(cp_addr1, "_", "") & ", " & replace(cp_addr2, "_", "")
		ELSE
			cp_addr = "Unknown Address"
		END IF
	END PROPERTY

	' CP address city
	PUBLIC PROPERTY GET cp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadscreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_city, 20, 17, 11
			cp_city = replace(cp_city, "_", "")
		ELSE
		cp_city = "City"
		END IF
	END PROPERTY

	' CP address state
	PUBLIC PROPERTY GET cp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadscreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_state, 2, 17, 39
		ELSE
			cp_state = "State"
		END IF
	END PROPERTY
	
    ' CP address zip code
	PUBLIC PROPERTY GET cp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen cp_zip, 10, 17, 50
		ELSE
		cp_zip = "ZIP"
		END IF
	END PROPERTY
	
	' CP employer
	PUBLIC PROPERTY GET cp_employer
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPSU" THEN CALL navigate_to_PRISM_screen("CPSU")
		EMReadScreen cp_employer, 30, 11, 12	
	END PROPERTY

	' CP phone numbers
	PUBLIC PROPERTY GET cp_phone_numbers
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_home_phone, 12, 13, 14
		cp_home_phone = Replace(cp_home_phone, "_", "")
		cp_home_phone = Replace(cp_home_phone, " ", "-")
		IF cp_home_phone <> "--" THEN	
			cp_phone_numbers = cp_phone_numbers & "Home: " & cp_home_phone & "; "
		ELSE
			cp_phone_numbers = cp_phone_numbers & "No known home phone number."
		END IF
		EMReadScreen cp_cell_phone, 12, 14, 14
		cp_cell_phone = Replace(cp_cell_phone, "_", "")
		cp_cell_phone = Replace(cp_cell_phone, " ", "-")
		IF cp_cell_phone <> "--" THEN
			cp_phone_numbers = cp_phone_numbers & " Cell: " & cp_cell_phone & "; "
		ELSE
			cp_phone_numbers = cp_phone_numbers & " No known cell phone number."
		END IF
		EMReadScreen cp_alt_phone, 12, 13, 40
		cp_alt_phone = Replace(cp_alt_phone, "_", "")
		cp_alt_phone = Replace(cp_alt_phone, " ", "-")
		IF cp_alt_phone <> "--" THEN
		
			cp_phone_numbers = cp_phone_numbers & " Alt: " & cp_alt_phone & "; "
		ELSE
			cp_phone_numbers = cp_phone_numbers & " No known alt phone number."
		END IF
			
	END PROPERTY

	' # of CP's open cases
	PUBLIC PROPERTY GET number_of_cps_open_cases
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPCB" THEN CALL navigate_to_PRISM_screen("CPCB")

		browse_row = 7
		open_cases = 0
		DO
			EMReadScreen end_of_data, 11, browse_row, 32
		
			if end_of_data <> "End of Data" then
		
				EMReadScreen browse_role, 5, browse_row, 8
				EMReadScreen browse_stat, 3, browse_row, 68
				EMReadScreen browse_case_num_first, 10, browse_row, 15
				EMReadScreen browse_case_num_second, 2, browse_row, 26

			'Check the role and case status - we only want active cases where participant is not the child
				If browse_role <> "Child" and browse_stat = "OPN" then
					open_cases = open_cases + 1
				END IF
				browse_row = browse_row + 1
				IF browse_row = 19 THEN
					PF8
					browse_row = 8
				END IF
			END IF
		LOOP UNTIL end_of_data = "End of Data"
		number_of_cps_open_cases = open_cases
	
	END PROPERTY

	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>>>> NCP Information <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' NCP Name
	PUBLIC PROPERTY GET ncp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_name, 50, 5, 25
		ncp_name = trim(ncp_name)
	END PROPERTY
	
	' NCP first name
	PUBLIC PROPERTY GET ncp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_first_name, 12, 8, 34
		ncp_first_name = trim(replace(ncp_first_name, "_", ""))
	END PROPERTY

	' NCP last name
	PUBLIC PROPERTY GET ncp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_last_name, 17, 8, 8
		ncp_last_name = trim(replace(ncp_last_name, "_", ""))
	END PROPERTY	
	
	' NCP middle name
	PUBLIC PROPERTY GET ncp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_middle_name, 12, 8, 56
		ncp_middle_name = trim(replace(ncp_middle_name, "_", ""))
	END PROPERTY
	
	' NCP middle initial
	PUBLIC PROPERTY GET ncp_middle_initial
		ncp_middle_initial = left(ncp_middle_name, 1)
	END PROPERTY
	
	' NCP suffix
	PUBLIC PROPERTY GET ncp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")	
		EMReadScreen ncp_suffix, 3, 8, 74
		ncp_suffix = trim(replace(ncp_suffix, "_", ""))
	END PROPERTY	
	
	' NCP date of birth
	PUBLIC PROPERTY GET ncp_dob
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_dob, 8, 6, 24		
	END PROPERTY

	' NCP SSN
	PUBLIC PROPERTY GET ncp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_ssn, 11, 6, 7
	END PROPERTY
	
	' NCP MCI
	PUBLIC PROPERTY GET ncp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_mci, 10, 5, 7
	END PROPERTY	

	' NCP street address
	PUBLIC PROPERTY GET ncp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_addr1, 30, 15, 11
			EMReadScreen ncp_addr2, 30, 16, 11
			ncp_addr = replace(ncp_addr1, "_", "") & ", " & replace(ncp_addr2, "_", "")
		ELSE
			ncp_addr = "Unknown Address"
		END IF
	END PROPERTY

	' NCP address city
	PUBLIC PROPERTY GET ncp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_city, 20, 17, 11
			ncp_city = replace(ncp_city, "_", "")
		ELSE
			ncp_city = "City"
		END IF
	END PROPERTY

	' NCP address state
	PUBLIC PROPERTY GET ncp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_state, 2, 17, 39
		ELSE
			ncp_state = "State"
		END IF
	END PROPERTY
    
	' NCP address zip code
	PUBLIC PROPERTY GET ncp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen known, 1, 10, 46
		IF known = "Y" THEN
			EMReadScreen ncp_zip, 10, 17, 50
		ELSE
			ncp_zip = "ZIP"
		END IF
	END PROPERTY

	' NCP employer
	PUBLIC PROPERTY GET ncp_employer
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCSU" THEN CALL navigate_to_PRISM_screen("NCSU")
		EMReadScreen ncp_employer, 30, 13, 49	
	END PROPERTY

	' NCP phone numbers
	PUBLIC PROPERTY GET ncp_phone_numbers
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_home_phone, 12, 13, 14
		ncp_home_phone = Replace(ncp_home_phone, "_", "")
		ncp_home_phone = Replace(ncp_home_phone, " ", "-")
		IF ncp_home_phone <> "--" THEN	
			ncp_phone_numbers = ncp_phone_numbers & "Home: " & ncp_home_phone & "; "
		ELSE
			ncp_phone_numbers = ncp_phone_numbers & "No known home phone number."
		END IF
		EMReadScreen ncp_cell_phone, 12, 14, 14
		ncp_cell_phone = Replace(ncp_cell_phone, "_", "")
		ncp_cell_phone = Replace(ncp_cell_phone, " ", "-")
		IF ncp_cell_phone <> "--" THEN
			ncp_phone_numbers = ncp_phone_numbers & " Cell: " & ncp_cell_phone & "; "
		ELSE
			ncp_phone_numbers = ncp_phone_numbers & " No known cell phone number."
		END IF
		EMReadScreen ncp_alt_phone, 12, 13, 40
		ncp_alt_phone = Replace(ncp_alt_phone, "_", "")
		ncp_alt_phone = Replace(ncp_alt_phone, " ", "-")
		IF ncp_alt_phone <> "--" THEN
			
			ncp_phone_numbers = ncp_phone_numbers & " Alt: " & ncp_alt_phone & "; "
		ELSE
			ncp_phone_numbers = ncp_phone_numbers & " No known alt phone number."
		END IF
		
	END PROPERTY

	' # of NCP's active cases
	PUBLIC PROPERTY GET number_of_ncps_open_cases
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCCB" THEN CALL navigate_to_PRISM_screen("NCCB")

		browse_row = 7
		active_case = false
		open_cases = 0
		DO
			EMReadScreen end_of_data, 11, browse_row, 32
		
			if end_of_data <> "End of Data" then
		
				EMReadScreen browse_role, 3, browse_row, 8
				EMReadScreen browse_stat, 3, browse_row, 68
				EMReadScreen browse_case_num_first, 10, browse_row, 15
				EMReadScreen browse_case_num_second, 2, browse_row, 26

			'Check the role and case status - we only want active cases where participant is not the child
				If browse_role <> "Child" and browse_stat = "OPN" then
					open_cases = open_cases + 1
				END IF
				browse_row = browse_row + 1
				IF browse_row = 19 THEN
					PF8
					browse_row = 8
				END IF
			END IF
		LOOP UNTIL end_of_data = "End of Data"
		number_of_ncps_open_cases = open_cases
	
	END PROPERTY

	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>> Financial Information <<<
	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<

	'basic support obligation
	PUBLIC PROPERTY GET cch_amount
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCOL" THEN CALL navigate_to_PRISM_screen("NCOL")
		EMWriteScreen "CCH", 20, 39
		transmit
		EMReadScreen cch_amount, 9, 9, 36
		IF inStr(Cstr(CCH_amount), "Data") > 0 THEN
			cch_amount = "0.00"
		ELSE
			cch_amount = trim(cch_amount)
		END IF
	END PROPERTY
	
	'child care support obligation
	PUBLIC PROPERTY GET ccc_amount
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCOL" THEN CALL navigate_to_PRISM_screen("NCOL")
		EMWriteScreen "CCC", 20, 39
		transmit
		EMReadScreen ccc_amount, 9, 9, 36
		IF inStr(Cstr(ccc_amount), "Data") > 0 THEN
			ccc_amount = "0.00"
		ELSE
			ccc_amount = trim(ccc_amount)
		END IF
	END PROPERTY

	'medical support obligation
	PUBLIC PROPERTY GET cms_amount
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCOL" THEN CALL navigate_to_PRISM_screen("NCOL")
		EMWriteScreen "CMS", 20, 39
		transmit
		EMReadScreen cms_amount, 9, 9, 36
		IF inStr(Cstr(cms_amount), "Data") > 0 THEN
			cms_amount = "0.00"
		ELSE
			cms_amount = trim(cms_amount)
		END IF
	END PROPERTY

	'medical insurance contribution 
	PUBLIC PROPERTY GET cmi_amount
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCOL" THEN CALL navigate_to_PRISM_screen("NCOL")
		EMWriteScreen "CMI", 20, 39
		transmit
		EMReadScreen cmi_amount, 9, 9, 36
		IF inStr(Cstr(cmi_amount), "Data") > 0 THEN
			cmi_amount = "0.00"
		ELSE
			cmi_amount = trim(cmi_amount)
		END IF
	END PROPERTY

	'spousal support obligation
	PUBLIC PROPERTY GET csp_amount
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCOL" THEN CALL navigate_to_PRISM_screen("NCOL")
		EMWriteScreen "CSP", 20, 39
		transmit
		EMReadScreen csp_amount, 9, 9, 36
		IF inStr(Cstr(csp_amount), "Data") > 0 THEN
			csp_amount = "0.00"
		ELSE
			csp_amount = trim(csp_amount)
		END IF
	END PROPERTY

	' monthly accrual amount
	PUBLIC PROPERTY GET monthly_accrual
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen monthly_accrual, 8, 9, 31
		monthly_accrual = trim(monthly_accrual)
	END PROPERTY
	
	' monthly non-accrual
	PUBLIC PROPERTY GET monthly_non_accrual
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen monthly_non_accrual, 8, 10, 31
		monthly_non_accrual = trim(monthly_non_accrual)
	END PROPERTY
	
	' NPA arrears
	PUBLIC PROPERTY GET npa_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen npa_arrears, 13, 10, 65
		npa_arrears = trim(npa_arrears)
	END PROPERTY
	
	' PA arrears
	PUBLIC PROPERTY GET pa_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen pa_arrears, 13, 11, 65
		pa_arrears = trim(pa_arrears)
	END PROPERTY
	
	' Total arrears
	PUBLIC PROPERTY GET ttl_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen ttl_arrears, 13, 12, 65
		ttl_arrears = trim(ttl_arrears)
	END PROPERTY

	' Last payment date
	PUBLIC PROPERTY GET last_payment_date
		EMReadScreen at_screen, 20, 2, 29
		IF at_screen <> "Payment List By Case" THEN CALL navigate_to_PRISM_screen("PALC")
		EMWritescreen "12/12/2015", 20, 49
		transmit
		EMWriteScreen date, 20, 49
		transmit
		EMWriteScreen "D", 9, 5
		transmit
		EMReadScreen end_of_data, 11, 9, 32
			if end_of_data = "End of Data" then
				last_payment_date = "N/A"
			else
				EMReadScreen last_payment_date, 8, 13, 37
			end if
	END PROPERTY

	'Last payment type
	PUBLIC PROPERTY GET last_payment_type
		EMReadScreen at_screen, 20, 2, 29
		IF at_screen <> "Payment List By Case" THEN CALL navigate_to_PRISM_screen("PALC")
		EMWritescreen "12/12/2015", 20, 49
		transmit
		EMWriteScreen date, 20, 49
		transmit
		EMReadScreen end_of_data, 11, 9, 32
			if end_of_data = "End of Data" then
				last_payment_type = "N/A"
			else
				EMReadScreen last_payment_type, 3, 9, 25
				 IF last_payment_type = "STJ" or last_payment_type = "STS" or last_payment_type = "FTJ" or last_payment_type = "FTS" or last_payment_type = "FIN"_
					 or last_payment_type = "BND" THEN last_payment_type = "Involuntary"
			end if
	END PROPERTY

	'Last payment amount
	PUBLIC PROPERTY GET last_payment_amount
		EMReadScreen at_screen, 20, 2, 29
		IF at_screen <> "Payment List By Case" THEN CALL navigate_to_PRISM_screen("PALC")
		EMWritescreen "12/12/2015", 20, 49
		transmit	
		EMWriteScreen date, 20, 49
		transmit
		EMReadScreen end_of_data, 11, 9, 32
			if end_of_data = "End of Data" then
				last_payment_amount = "0.00"
			else
				EMReadScreen last_payment_amount, 13, 9, 29
				last_payment_amount = trim(last_payment_amount)
			end if
	END PROPERTY

	'Last payment allocation
	PUBLIC PROPERTY GET last_payment_allocation
		EMReadScreen at_screen, 20, 2, 29
		IF at_screen <> "Payment List By Case" THEN CALL navigate_to_PRISM_screen("PALC")
		EMWritescreen "12/12/2015", 20, 49
		transmit
		EMWriteScreen date, 20, 49
		transmit
		EMReadScreen end_of_data, 11, 9, 32
			if end_of_data = "End of Data" then
				last_payment_allocation = "0.00"
			else
				EMReadScreen last_payment_allocation, 12, 9, 68
				last_payment_allocation = trim (last_payment_allocation)
			end if
	END PROPERTY
	
	'Last payment plan info
	PUBLIC PROPERTY GET pay_plan_info(case_number)
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "PAPD" THEN CALL navigate_to_PRISM_screen("PAPD")
		EMWritescreen "B", 3, 29
		transmit

		EMWriteScreen Left(case_number, 10), 19, 25
		EMWriteScreen Right(case_number, 2), 19, 36
		EMWriteScreen "DLS", 4, 48 
		transmit
		
		EMReadScreen end_of_data, 11, 8, 32
		IF end_of_data = "End of Data" THEN
			pay_plan_info = "No DL pay plan information found."
		ELSE
			EMReadScreen pay_plan_begin, 8, 8, 47
			EMReadScreen pay_plan_end, 8, 8, 57
		
			EMSetCursor 8, 32
			transmit
		
			EMReadScreen delinquent_amt, 12, 15, 17
			delinquent_amt = trim(delinquent_amt)
			pay_plan_info = "Begin date: " & pay_plan_begin & "; End date: " & pay_plan_end & "; Amount Delinquent: " & FormatCurrency(delinquent_amt)
		END IF

	END PROPERTY


	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>> General Information <<<
	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<

	' Case worker name Last, First
	PUBLIC PROPERTY GET worker_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMSetCursor 5, 56
		PF1
		EMReadScreen worker_name, 30, 6, 50
		worker_name = trim(worker_name)
		transmit
	END PROPERTY
	
	' Case worker phone ###-###-####
	PUBLIC PROPERTY GET worker_phone
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMSetCursor 5, 56
		PF1
		EMReadScreen worker_phone, 12, 8, 35
		transmit
	END PROPERTY	

	' Case function
	PUBLIC PROPERTY GET case_function
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen case_function, 2, 5, 78
	END PROPERTY	
END CLASS
'------End of Class definitions--------------------------------------------------------------------------------

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


'Displays the Interview Info dialog so the user can pick which participant to create the document for.
Dialog Interview_Info_dialog
If buttonpressed = 0 then stopscript

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
call navigate_to_PRISM_screen("CAPS")

'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it

'The command below is necessary in order to utilize the doc_info class to efficiently obtain case data.  We are creating an object called "info" that is a member of the "doc_info" class.
set info = new doc_info

'Create a Microsoft Word object, make it visible
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

'This script is designed to create a Word document for an in-person meeting with a client.  The top portion displays client information with space for the worker to mark for PRISM updates.
'The bottom part of the document is intended to be a list of resources for the client, and can be folded in half and torn from the top half to be given to the client.
'The Word template is able to be modified by individual counties so it can reference local resources.  As agencies make changes, be sure not to delete any fields.
'Agencies should have their template for this document saved as "interview-information-sheet.docx" and saved in the agencies' Word document file path location set in global variables.  The same template is used 
'whether the user indicates it is produced for CP or NCP.  However, different data appears in the fields depending on whether the document is produced for CP or NCP.


'If the participant selected by the user is NCP, then the script completes the form fields for NCP using information produced by using the doc_info class methods. 
'On line 731, we named our doc_info class object, "info". To efficiently obtain case information, we call methods for the property information of our doc_info object.
'To call, start with "info", then a period, and then the name of the property.  
'To review the properties you can use, check out the class definition in lines 69-684.
'Unless classes are added to the functions library, it is necessary to have the class definition in your script if you want to use it.
IF participant = "NCP" Then	
	set objDoc = objWord.Documents.Add(word_documents_folder_path & "interview-information-sheet.docx")
		With objDoc
			.FormFields("user").Result = worker_signature
			.FormFields("name").Result = info.ncp_name   
			.FormFields("address").Result = info.ncp_addr & " " & info.ncp_city & ", " & info.ncp_state & " " & Left(info.ncp_zip, 5)
			.FormFields("phone").Result = info.ncp_phone_numbers
			.FormFields("employer").Result = info.ncp_employer
			.FormFields("mci").Result = info.ncp_mci	
			.FormFields("case_number").Result = PRISM_case_number
			.FormFields("other_cases").Result = info.number_of_ncps_open_cases		
			.FormFields("function").Result = info.case_function
			.FormFields("assigned_worker").Result = info.worker_name & ", Phone: " & info.worker_phone
			.FormFields("last_payment").Result = info.last_payment_date & " " & FormatCurrency(CCur(info.last_payment_allocation)) & " " & info.last_payment_type
			.FormFields("arrears_balance").Result = FormatCurrency (info.pa_arrears) & " PA arrears + " & FormatCurrency(info.npa_arrears) & " NPA arrears = " & FormatCurrency(info.ttl_arrears) & " total"
			.FormFields("pay_plan").Result = info.pay_plan_info(PRISM_case_number)
			'The dollar values below must be converted to currency using the CCur before we can use the FormatCurrency fucntion.  The end result is beautifully formatted dollar amounts with minimal effort!
			.FormFields("cch").Result = FormatCurrency (Ccur(info.cch_amount))
			.FormFields("cms").Result = FormatCurrency(CCur(info.cms_amount))
			.FormFields("cmi").Result = FormatCurrency(CCur(info.cmi_amount))
			.FormFields("ccc").Result = FormatCurrency(CCur(info.ccc_amount))
			.FormFields("csp").Result = FormatCurrency(CCur(info.csp_amount))
			.FormFields("arrears_payback").Result = FormatCurrency(info.monthly_non_accrual)
			.FormFields("ncp_mci").Result = info.ncp_mci	
	End With
ELSE
'The participant selected by the user is CP.  The script completes the form fields with CP's information using doc_info class's methods. 
'On line 731, we named our doc_info class object, "info". To efficiently obtain case information, we call methods for the property information of our doc_info object.
'To call, start with "info", then a period, and then the name of the property.  
'To review the properties you can use, check out the class definition in lines 69-684.
'Unless classes are added to the functions library, it is necessary to have the class definition in your script if you want to use it.
	set objDoc = objWord.Documents.Add(word_documents_folder_path & "interview-information-sheet.docx")
		With objDoc
			.FormFields("user").Result = worker_signature
			.FormFields("name").Result = info.cp_name
			.FormFields("address").Result = info.cp_addr & " " & info.cp_city & ", " & info.cp_state & " " & Left(info.cp_zip, 5)
			.FormFields("phone").Result = info.cp_phone_numbers
			.FormFields("employer").Result = info.cp_employer 
			.FormFields("mci").Result = info.cp_mci	
			.FormFields("case_number").Result = PRISM_case_number
			.FormFields("other_cases").Result = info.number_of_cps_open_cases	
			.FormFields("function").Result = info.case_function
			.FormFields("assigned_worker").Result = info.worker_name & ", Phone: " & info.worker_phone
			.FormFields("last_payment").Result = info.last_payment_date & " " & FormatCurrency(info.last_payment_allocation) & " " & info.last_payment_type	
			.FormFields("arrears_balance").Result = FormatCurrency (info.pa_arrears) & " PA arrears + " & FormatCurrency(info.npa_arrears) & " NPA arrears = " & FormatCurrency(info.ttl_arrears) & " total"
			.FormFields("pay_plan").Result = info.pay_plan_info(PRISM_case_number)
			'The dollar values below must be converted to currency using the CCur before we can use the FormatCurrency fucntion.  The end result is beautifully formatted dollar amounts with minimal effort!
			.FormFields("cch").Result = FormatCurrency (Ccur(info.cch_amount))
			.FormFields("cms").Result = FormatCurrency(CCur(info.cms_amount))
			.FormFields("cmi").Result = FormatCurrency(CCur(info.cmi_amount))
			.FormFields("ccc").Result = FormatCurrency(CCur(info.ccc_amount))
			.FormFields("csp").Result = FormatCurrency(CCur(info.csp_amount))
			.FormFields("arrears_payback").Result = FormatCurrency(info.monthly_non_accrual)
			.FormFields("ncp_mci").Result = info.ncp_mci	
	End With
END IF
script_end_procedure("")
