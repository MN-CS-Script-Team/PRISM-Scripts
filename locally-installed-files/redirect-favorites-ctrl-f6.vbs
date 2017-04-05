'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
GlobVar_path = "C:\DHS-PRISM-Scripts\locally-installed-files\~globvar.vbs"													'Setting a default path, which is modified by the installer
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")														'Creating an FSO for the work
If run_another_script_fso.FileExists(GlobVar_path) then																		'If a Global Variables file is found in above directory...
	Set fso_command = run_another_script_fso.OpenTextFile(GlobVar_path)														'...run it!
Else																														'If a Global Variables file is not found in the above directory...
	Set fso_command = run_another_script_fso.OpenTextFile("Scripts\~globvar-local.vbs")										'...use the default BlueZone Scripts directory, and insert a custom "local flavor" Global Variables file, which can override the default selections.
End if
text_from_the_other_script = fso_command.ReadAll																			'Once we have the text from the other script, read it all!
fso_command.Close																											'Close the other script file, and...
Execute text_from_the_other_script

'LOADING SCRIPT
script_URL = script_repository & "/favorites/ctrl-f6.vbs"
IF run_locally = False THEN
	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
	req.open "GET", script_URL, FALSE									'Attempts to open the URL
	req.send													'Sends request
	IF req.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
		Execute req.responseText								'Executes the script code
	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
										"Script URL: " & script_URL & vbNewLine & vbNewLine &_
										"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
										vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
		StopScript
	END IF
ELSE
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(script_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF
