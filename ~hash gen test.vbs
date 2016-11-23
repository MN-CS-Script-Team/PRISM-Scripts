function md5hashBytes(aBytes)
    Dim MD5
    set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    MD5.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    md5hashBytes = MD5.ComputeHash_2( (aBytes) )
end function

function sha1hashBytes(aBytes)
    Dim sha1
    set sha1 = CreateObject("System.Security.Cryptography.SHA1Managed")

    sha1.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha1hashBytes = sha1.ComputeHash_2( (aBytes) )
end function

function sha256hashBytes(aBytes)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    sha256.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha256hashBytes = sha256.ComputeHash_2( (aBytes) )
end function

function stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
end function

function bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        bytesToHex=bytesToHex & hexStr
    next
end function

Function BytesToBase64(varBytes)
    With CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = .Text
    End With
End Function

Function GetBytes(sPath)
    With CreateObject("Adodb.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .LoadFromFile sPath
        .Position = 0
        GetBytes = .Read
        .Close
    End With
End Function

'<<<<TEMP VARIABLES, I'll be turning this into a batch process or something once I know where I can store the hashes online (SIR, CountyLink, separate server?)
'Do we want to generate a hash, or simply compare? Set to "true" to build a new hash for testing purposes after a change is made.
generate = False

'Needs to determine MyDocs directory before proceeding. Used purely to find the local hash path for testing. Eventually this will be hard-coded into a global variables file.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

'URL we want to work through. For now, we're just picking a CAAD navigation script (very straight forward navigation script)
'Sets the URL (this will likely be passed as a parameter in a JS function or even using GitHub's API and an OAuth hook)
script_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/nav/caad.vbs"

'Name of the local file we're storing the hash in. For now it's a single file for a single script, obviously we'll do more with this once server details are resolved
local_hash_file_path = user_myDocs_folder & "CAAD-hash.txt"

'=============================THE ACTUAL FUNCTIONALITY

'If generate is true, it'll create a new hash file from my local version (this is purely for testing, I will deploy a JavaScript hook for auto generation on a schedule of some kind.
If generate = true then

	'Grabs the script file from GitHub's raw user content (this could just as easily be a different server, but would require every county to reinstall)
	SET req = CreateObject("Msxml2.XMLHttp.6.0")						'Creates an object to get a URL
	req.open "GET", script_URL, FALSE								'Attempts to open the URL
	req.send												'Sends request
	IF req.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
		text_from_GitHub_script = req.responseText					'Reads the script code into a variable
	ELSE													'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
										"Script URL: " & script_URL & vbNewLine & vbNewLine &_
										"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
										vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
		StopScript
	END IF

	'Opens an FSO, opens CAAD-hash.txt, writes the new hash in, and exits
	SET update_CAAD_hash_fso = CreateObject("Scripting.FileSystemObject")
	SET update_CAAD_hash_command = update_CAAD_hash_fso.CreateTextFile(local_hash_file_path, 2)
	update_CAAD_hash_command.Write(bytesToHex(sha256hashBytes(stringToUTFBytes(text_from_GitHub_script))))
	update_CAAD_hash_command.Close
Else

	'Grabs the script file from GitHub's raw user content (this could just as easily be a different server, but would require every county to reinstall)
	SET req = CreateObject("Msxml2.XMLHttp.6.0")						'Creates an object to get a URL
	req.open "GET", script_URL, FALSE								'Attempts to open the URL
	req.send												'Sends request
	IF req.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")				'Creates an FSO
		text_from_GitHub_script = req.responseText					'Reads the script code into a variable
	ELSE													'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
										"Script URL: " & script_URL & vbNewLine & vbNewLine &_
										"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
										vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
		StopScript
	END IF
End if

'If generate = false (in the future, the normal behavior), let's compare the two versions of the hash! If it doesn't match (i.e. something was changed) it'll force an alert
If generate = false then
	'An SHA-256 hash from the file on GitHub
	hashed_version_of_file = bytesToHex(sha256hashBytes(stringToUTFBytes(text_from_GitHub_script)))

	'Open the hash that was previously saved to the computer or local network
	SET update_CAAD_hash_fso = CreateObject("Scripting.FileSystemObject")
	SET update_CAAD_hash_command = update_CAAD_hash_fso.OpenTextFile(local_hash_file_path)
	stored_hash = update_CAAD_hash_command.ReadAll
	update_CAAD_hash_command.Close

	'Compare them!
	If stored_hash = hashed_version_of_file then
		MsgBox "Files match! You're good to go!"
	Else
		MsgBox "Warning! File doesn't match locally stored hash!"
	End if
end if