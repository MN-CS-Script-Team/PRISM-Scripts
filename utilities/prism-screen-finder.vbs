'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "prism-screen-finder.vbs"
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
call changelog_update("11/23/2016", "The script has been moved from the NAV category to the UTILILTIES category. A new Power Pad will be released, which will remove this script (and add long-awaited FAVORITES functionality).", "Veronica Cary, DHS")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


BeginDialog Find_that_screen_in_PRISM_dialog, 0, 0, 161, 185, "Find that Screen in PRISM"
  ButtonGroup ButtonPressed
    OkButton 45, 160, 50, 15
    CancelButton 100, 160, 50, 15
    PushButton 115, 5, 35, 10, "ACSD", ACSD_button
    PushButton 115, 20, 35, 10, "CPRE", CPRE_button
    PushButton 115, 35, 35, 10, "GCSC", GCSC_button
    PushButton 115, 50, 35, 10, "NCLD", NCLD_button
    PushButton 115, 65, 35, 10, "NCLL", NCLL_button
    PushButton 115, 80, 35, 10, "NCSL", NCSL_button
    PushButton 115, 95, 35, 10, "PALI", PALI_button
    PushButton 115, 110, 35, 10, "REID", REID_button
    PushButton 115, 125, 35, 10, "SEPD", SEPD_button
    PushButton 115, 140, 35, 10, "WEDL", WEDL_button
  Text 10, 5, 70, 10, "Account status detail"
  Text 10, 20, 70, 10, "CP/NCP Relationship"
  Text 10, 35, 70, 10, "Good Cause Screen"
  Text 10, 50, 80, 10, "NCP license data detail"
  Text 10, 65, 70, 10, "NCP license data list"
  Text 10, 80, 70, 10, "NCP alias detail"
  Text 10, 95, 70, 10, "Payment listing"
  Text 10, 110, 90, 10, "Re-employment ins detail"
  Text 10, 125, 90, 10, "Service of Process detail"
  Text 10, 140, 95, 10, "On-line employer reporting"
EndDialog

'The Script--------------------------------------------------------------------------------------------------

Dialog Find_that_screen_in_PRISM_dialog

'Connect to BlueZone
EMConnect ""

CALL check_for_PRISM(true)			'this checks whether in PRISM or timed out

'Not adding case number function as PRISM will go to screen without case number or PRISM populates with last case number used.

'Naviagting to any of the buttons chosen
If buttonpressed = ACSD_button then call navigate_to_PRISM_screen("ACSD")
If buttonpressed = CPRE_button then call navigate_to_PRISM_screen("CPRE")
If buttonpressed = GCSC_button then call navigate_to_PRISM_screen("GCSC")
If buttonpressed = NCLD_button then call navigate_to_PRISM_screen("NCLD")
If buttonpressed = NCLL_button then call navigate_to_PRISM_screen("NCLL")
If buttonpressed = NCSL_button then call navigate_to_PRISM_screen("NCSL")
If buttonpressed = PALI_button then call navigate_to_PRISM_screen("PALI")
If buttonpressed = REID_button then call navigate_to_PRISM_screen("REID")
If buttonpressed = SEPD_button then call navigate_to_PRISM_screen("SEPD")
If buttonpressed = WEDL_button then call navigate_to_PRISM_screen("WEDL")

script_end_procedure("")
