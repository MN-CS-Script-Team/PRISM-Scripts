'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "BULK - M1600 WORKLIST SCRUBBER.vbs" 
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

' >>>>> THE SCRIPT <<<<<
EMConnect ""

CALL select_cso_caseload(ButtonPressed, cso_id, cso_name)	'LETS YOU SELECT WHICH CASE NUMBER YOU WOULD LIKE TO RUN THE SCRIPT ON


count = 0
USWT_row = 7


Call navigate_to_Prism_screen("USWT")				'BRINGS YOU TO USWT


EMWriteScreen "M1600", 20, 30						'SELECTING THE SPECIFIC WORKLIST TYPE INTO CAWT
transmit									'ENTER

DO
	EMReadScreen USWT_type, 5, USWT_row, 45 			'NEED TO FILTER TO THE WORKLIST THAT WE ARE WORKING WITH 
	IF USWT_type = "M1600" THEN
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		EMWriteScreen "S", USWT_row, 4			'SELECTS THE WORKLISTS AND BRINGS YOU TO THE CAST SCREEN
		transmit							'ENTERS WHAT WAS ENTERED ABOVE
		EMReadScreen Closure_Reas_check, 3, 13, 52 	'READS THE CLOSURE REASON
		If Closure_Reas_check = "901" or Closure_Reas_check = "902" or Closure_Reas_check = "910" or Closure_Reas_check = "911" or Closure_Reas_check = "916" or Closure_Reas_check = "923" or Closure_Reas_check = "940" or Closure_Reas_check = "950" THEN
			Call navigate_to_Prism_screen ("CAWT")	'BRINGS YOU TO THE CAWT SCREEN IN THE CASE THAT HAS THE ABOVE WORKLISTS
			EMWriteScreen "M1600", 20, 29			'BRIINGS M1600 TO THE TOP SO THAT IT CAN SELECT THE WORKLIST
			transmit						
			EMReadScreen CAWT_type, 5, 8, 8		
			IF CAWT_type = "M1600" THEN    		'PURGES M1600 IF IT IS ONE OF THE ABOVE REASON TYPES
				EMWriteScreen "P", 8, 4
				transmit
				transmit
				count = count + 1      			'COUNTS THE NUMBER OF WORKLISTS THAT HAVE BEEN PURGED
 			END IF 
		ELSE
			USWT_row = USWT_row + 1				'IF WORKLIST WAS NOT PURGE THEN WILL BRING YOU TO THE NEXT LINE TO VIEW THE NEXT WORKLIST 
		End if
	END IF

	
	Call navigate_to_Prism_screen("USWT")			'BRINGING US BACK TO USWT SO CAN SELECT THE NEXT WORKLIST 

	EMWriteScreen "M1600", 20, 30					'SELECTING THE WORKLISTS AGAIN
	transmit

LOOP Until USWT_type <> "M1600"					'RERUNS THE SCRIPT UNTIL ALL WORKLISTS HAVE BEEN REVIEWED IN THAT CASE NUMBER

script_end_procedure("Success!  " & count & " worklists were purged.") 		'TELLS YOU HOW MANY WORKLISTS WERE PURGED BY RUNNING THIS SCRIPT
