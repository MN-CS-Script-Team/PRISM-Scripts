'GATHERING STATS=================================
name_of_script = "free-worklist.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	                        'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			                  'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE						'Attempts to open the FuncLib_URL
		req.send										'Sends request
		IF req.Status = 200 THEN							'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	      'Creates an FSO
			Execute req.responseText						'Executes the script code
		ELSE											'Error message
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
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE DIALOG BOX-------------------------------------------------------------------------------------------------------------------
BeginDialog free_worklist_dialog, 0, 0, 186, 145, "Free Worklist"
  Text 5, 10, 50, 10, "Case number:"
  EditBox 65, 5, 75, 15, PRISM_case_number
  Text 5, 25, 20, 10, "Type"
  DropListBox 65, 25, 60, 15, "A Action needed"+chr(9)+"I Information Only", ref_droplistbox
  Text 5, 45, 30, 10, "Category"
  DropListBox 65, 40, 60, 15, "Base Case"+chr(9)+"CAO Response"+chr(9)+"Check Odyssey"+chr(9)+"Close Case"+chr(9)+"EMC Child"+chr(9)+"Maintaining County"+chr(9)+"Medical Only Case"+chr(9)+"Other"+chr(9)+"Probation Officer"+chr(9)+"Redirection Case"+chr(9)+"Refund FREM"+chr(9)+"Release on File"+chr(9)+"Request/Retrieve Acre"+chr(9)+"Response from Party"+chr(9)+"Social Worker"+chr(9)+"Work Number Response?"+chr(9)+"Zero/Reserved Order"+chr(9)+"##Interstate", ref_droplistbox
  Text 5, 60, 50, 10, "Additional Info"
  EditBox 70, 60, 95, 15, Additional_Info
  Text 5, 85, 60, 10, "Follow Up Needed"
  EditBox 70, 80, 95, 15, Follow_Up_Needed
  Text 5, 110, 95, 10, "Last Reviewed Reviewed By"
  EditBox 105, 105, 60, 15, Last_Reviewed_Reviewed_By
  ButtonGroup ButtonPressed
    OkButton 65, 125, 50, 15
    CancelButton 125, 125, 50, 15
EndDialog

'DIM row, col, EMSearch, EMReadScreen

'THE SCRIPT CODE-------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)


'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

Do
	err_msg = ""
	'Shows dialog, validates that PRISM is up and not timed out, with transmit
	Dialog case_initiation_docs_recd_dialog
	If buttonpressed = 0 then stopscript
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = False THEN err_msg = err_msg & vbNewLine & "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX.  "
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "Sign your CAAD note."
	IF err_msg <> "" THEN
				MsgBox "***NOTICE***" & vbcr & err_msg & vbNewline & vbNewline & "Please resolve for this script to continue."
	END IF
LOOP UNTIL err_msg = ""



CALL check_for_PRISM(True)                                      'Makes sure you are not passworded out


DO                                                              'MAKES THINGS MANDATORY

	err_msg = ""
	Dialog free_worklist_dialog
	cancel_confirmation
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF ref_droplistbox = "Select One..." THEN err_msg = err_msg & vbNewline & "You must select type!"
	IF ref_droplistbox = "Select One..." THEN err_msg = err_msg & vbNewline & "You must select category!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""


CALL navigate_to_PRISM_screen("CAWT")                           'Navigates to CAWT and adds the worklist


PF5                                                             'Adds new CAWT worklist


EMWriteScreen "A", 3, 30                                        'Writes the CAWT worklist



EMWriteScreen "FREE", 4, 37                                     'Type of CAWT worklist
EMWriteScreen 10, 4                                             'Types "*A:*" or "*I:*" ref_droplistbox and "Base Case"+chr(9)+"CAO Response"+chr(9)+"Check Odyssey"+chr(9)+"Close Case"+chr(9)+"EMC Child"+chr(9)+"Maintaining County"+chr(9)+"Medical Only Case"+chr(9)+"Other"+chr(9)+"Probation Officer"+chr(9)+"Redirection Case"+chr(9)+"Refund FREM"+chr(9)+"Release on File"+chr(9)+"Request/Retrieve Acre"+chr(9)+"Response from Party"+chr(9)+"Social Worker"+chr(9)+"Work Number Response?"+chr(9)+"Zero/Reserved Order"+chr(9)+"##Interstate", ref_droplistbox on first line of CAAD note
EMWriteScreen 10, 7                                             'Types Additional_Info
EMWriteScreen 11, 4                                             'Types Follow_Up_Needed
EMWriteScreen 13, 4                                             'Types Last_Reviewed_Reviewed_by



transmit                                                        'Saves the CAWT worklist


PF3                                                             'Exits back out of that CAWT worklist


script_end_procedure("")          'Stops the script
