'Gathering stats+=====================
'name_of_script = "ACTIONS - ESTB DORD DOCS FOR PA CASE.vbs"
'start_time = timer



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



'THIS SCRIPT IS READY FOR USE 


'This is an updated version of the ESTB PA DORD DOCS that is used when starting a NEW ESTABLISH ACTION on a public assistance case that prints the
'the financial statements, waivers, Important Statement of Rights and NCP Notice of Parental Liablity. It would NOT be used on a RELATIVE CARETAKER case as
'the CP is not required to complete financial docs on that type of case.

'Connecting to BZ  'This is a script for a PA case to print the fin docs & waivers
EMConnect ""   

'Checks to make sure we are in Prism
CALL check_for_Prism (true)

'Directing to DORD screen
Call navigate_to_Prism_Screen ("DORD")

'Clears the screen to add the doc
EMWriteScreen "C", 3,29

transmit

EMWriteScreen "A", 3,29

EMSetCursor 6,36
		
'adding the financial statement to DORD
EMWriteScreen "F0021", 6,36

'Printing financial statement x 2
transmit

pf9

transmit

pf9

transmit

EMSetCursor 3,29

EMWriteScreen "C", 3,29

transmit

'Adding the Important Statement of Rights
EMWriteScreen "A", 3,29

EMSetCursor 6,36			

EMWriteScreen "F0022", 6,36

'Printing the Important Statement of Rights x 2
Transmit

pf9

transmit

pf9

transmit

EMSetCursor 3,29

'Clearing screen for next doc
EMWriteScreen "C", 3,29

transmit

'Adding CP Waiver
EMWriteScreen "A", 3,29

EMSetCursor 6,36		

EMWriteScreen "F5000", 6,36

EMSetCursor 11,51

'Changing recipient to CP in DORD
EMWriteScreen "CPP", 11,51

transmit

EMSetCursor 3,29

'Modifying label in DORD
EMWriteScreen "M", 3,29	   
				    
pf14

pf8

EMSetCursor 13,5

'Selecting label line
EMWriteScreen "S", 13,5

transmit

'updating Waiver to say 12 months valid
EMWriteScreen "12", 16,15

transmit

pf3

'printing CP Waiver
pf9

transmit

EMSetCursor 3,29

'Clearing screen for next doc
EMWriteScreen "C", 3,29

transmit

'Adding Waiver to DORD
EmWriteScreen "A", 3,29

EMSetCursor 6,36

EMWriteScreen "F5000", 6,36

EMSetCursor 11,51		

'Changing recipient to NCP on Waiver
EMWriteScreen "NCP", 11,51

transmit

EMSetCursor 3,29

'Modifying label on DORD doc
EMWriteScreen "M", 3,29

pf14

pf8

EMSetCursor 13,5

'Selecting label line
EMWriteScreen "S", 13,5

transmit

'Modifying label to say Waiver valid for 12 months
EMWriteScreen "12", 16,15

transmit

pf3

'Pinting NCP Waiver
pf9

transmit	

EMSetCursor 3,29

'Clearing DORD screen
EMWriteScreen "C", 3,29

transmit
'Adding the NCP Notice of Liability
EMWriteScreen "A", 3,29   

EMSetCursor 6,36

EmwriteScreen "F0109", 6,36

transmit

EMWriteScreen "M", 3,29

pf14

EMSetCursor 20,14

EMWriteScreen "U", 20,14

transmit

'Selecting label line to include financial statement language on DORD doc
EMSetCursor 7,5

EMWriteScreen "S", 7,5

transmit

'Selecting the "Include Financial Statement" line
EMwriteScreen "X", 16,15 

transmit


'The Dialog to add worker information in the labels

DIM pa_dord_docs_dialog, worker_name_dialog, worker_title_dialog, worker_phone_dialog, ButtonPressed, write_variable_in_DORD
'INSERTED THE WORKER INFORMATION NEW DIALOG HERE

BeginDialog pa_dord_docs_dialog, 0, 0, 191, 135, "Worker Information Dialog"
  Text 10, 10, 50, 10, "Worker Name:"
  Text 10, 35, 45, 10, "Worker Title:"
  Text 10, 60, 55, 10, "Worker Phone:"
  EditBox 60, 5, 115, 15, worker_name_dialog
  EditBox 55, 30, 120, 15, worker_title_dialog
  EditBox 65, 55, 110, 15, worker_phone_dialog
  ButtonGroup ButtonPressed
    OkButton 10, 90, 50, 15
    CancelButton 10, 110, 50, 15
EndDialog

'This makes the dialog run
Dialog pa_dord_docs_dialog  

IF ButtonPressed = 0 THEN StopScript

EMSetCursor 9,5

EMWriteScreen "S", 9,5

transmit     'This next part below is needed to make the dialog box run in the script when entering the info

'Below writes the worker information typed into the DORD doc
EMWriteScreen (worker_name_dialog), 16,15 

transmit

EMSetCursor 10,5
EMWriteScreen "S", 10,5

transmit

EMWriteScreen (worker_title_dialog), 16,15  

transmit

EMSetCursor 11,5

EMWriteScreen "S", 11,5

transmit

EMWriteScreen (worker_phone_dialog), 16,15  

transmit

pf3

pf9

transmit 

'Adding the CAAD note
CALL navigate_to_PRISM_screen ("CAAD")

pf5
	
EMSetCursor 4,54    

EMWriteScreen "FREE", 4,54

EMSetCursor 16,4

EMWriteScreen "Sent CP and NCP Financial Statements and Waivers", 16,4

transmit

'Going to CAWT to write the tracking of the forms
Call navigate_to_Prism_Screen ("CAWT")

pf5

EMSetCursor 3,3

EMWriteScreen "A", 3,3

EMSetCursor 4,37   

EMWriteScreen "Free", 4,37

EMSetCursor 10,4


EMWriteScreen "Did CP & NCP Return Financial Statements and Waivers?", 10,4

EMSetCursor 17,52

'Writing the CAWT note for 14 days out
EMWriteScreen "14", 17,52

transmit

pf3

script_end_procedure("")

