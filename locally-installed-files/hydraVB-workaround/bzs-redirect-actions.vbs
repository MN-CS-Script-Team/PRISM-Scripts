' Determine Engine and architecture (this block re-runs scripts in cscript x86 if not initially started that way)
Engine = replace(replace(lcase(WScript.FullName), lcase(wscript.path), ""), "\", "")
If instr(lcase(wscript.path), "system32") then
    Architecture = "x64"
Elseif instr(lcase(wscript.path), "syswow64") then
    Architecture = "x86"
End if

' Create the complete command line to rerun this script in CSCRIPT32, and exit, if not already running in x86 cscript.exe.
If Engine <> "cscript.exe" or Architecture <> "x86" Then
    Set wshShell = CreateObject("WScript.Shell")
    set executecode = wshShell.exec("C:\Windows\SysWOW64\CSCRIPT.EXE //NoLogo """ & WScript.ScriptFullName & """")
    stdout_return = executecode.stdout.ReadAll
    stderr_return = executecode.stderr.ReadAll
'    MsgBox stdout_return

    if stderr_return <> "" then
        logfile_yn = MsgBox(stderr_return & vbNewLine & vbNewLine & _
                            "Your script will now stop. Would you like to save a log file?", _
                            vbYesNo, "Critical stop: an error was detected in the script")
        if logfile_yn = vbYes then

            logfile_location = wshShell.SpecialFolders("MyDocuments") & "\bzserr.txt"
            SET update_logfile_fso = CreateObject("Scripting.FileSystemObject")
            SET update_logfile_command = update_logfile_fso.CreateTextFile(logfile_location, 2)
            update_logfile_command.Write(stderr_return & vbNewLine & vbNewLine & stdout_return)
            update_logfile_command.Close
            MsgBox "Logfile saved to " & logfile_location

        End if
    end if

    ' Wait until the script exits
    Do While executecode.Status = 0
        WScript.Sleep 100
    Loop

    wscript.quit
end if



' LOADING GLOBAL VARIABLES
GlobVar_path = "bzs-global-variables.vbs"									'Setting a default path, which is modified by the installer
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")		'Creating an FSO for the work
Set fso_command = run_another_script_fso.OpenTextFile(GlobVar_path)			'...open it!
text_from_the_other_script = fso_command.ReadAll							'Once we have the text from the other script, read it all!
fso_command.Close								       						'Close the other script file, and...
ExecuteGlobal text_from_the_other_script



' LOADING HELPER FUNCTIONS FROM GITHUB
redirect_calling_github("https://raw.githubusercontent.com/MN-Script-Team/hydra/master/vbs-libs/bzio-helper-functions.vbs")

' LOADING THE ACTUAL MENU
redirect_calling_github("https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/~actions-menu-hydravb.vbs")



' FUNCTIONS ------------------------------------------------------------------------------------



function redirect_calling_github(script_URL)
 	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
 	req.open "GET", script_URL, FALSE									'Attempts to open the URL
 	req.send													'Sends request
 	IF req.Status = 200 THEN									'200 means great success
 		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
 		ExecuteGlobal req.responseText								'Executes the script code
 	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
 		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
 										"Script URL: " & script_URL & vbNewLine & vbNewLine &_
 										"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
 										vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
 		StopScript
 	END IF
end function
