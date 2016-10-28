'gathering stats===================
name_of_script = "NAV - EXTERNAL RESOURCES.vbs"
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

'This function builds the menu
FUNCTION external_resources_menu(default_directory, admin_enabled)
	'>>> READING THE TEXT FILE THAT CONTAINS THE LIST OF EXTERNAL RESOURCES <<<
	DIM oTxtFile
	WITH (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(default_directory & "\External Resources.txt") Then
			Set external_resources_file = CreateObject("Scripting.FileSystemObject")
			Set external_resources_file_command = external_resources_file.OpenTextFile(default_directory & "\External Resources.txt")		
			external_resources_array = external_resources_file_command.ReadAll
			IF external_resources_array = "" THEN 
				no_file_exists = MsgBox ("A resource file is required for this script to run. Press OK to continue to provide information on resources for your agency. Press CANCEL to stop this script.", vbOKCancel)
				IF no_file_exists = vbCancel THEN script_end_procedure("Script cancelled.")
				CALL create_external_resource(default_directory)
				script_end_procedure("")
			ELSE
				external_resources_file_command.Close
			END IF
		ELSE 
			no_file_exists = MsgBox ("A resource file is required for this script to run. Press OK to continue to provide information on resources for your agency. Press CANCEL to stop this script.", vbOKCancel)
			IF no_file_exists = vbCancel THEN script_end_procedure("Script cancelled.")
			CALL create_external_resource(default_directory)
		END IF			
	END WITH	

	'splitting the contents of the text file into an array
	external_resources_array = split(external_resources_array, vbNewLine)
	ReDim ext_res_multi_array(0, 5)
		'positions within the array
		'0 = name
		'1 = url
		'2 = description 
		'3 = pushbutton
		'4 = update button
		'5 = delete button
	
	'this is how we are going to determine ubound...there is an issue with the script creating too many vbNewLine s 
	array_position = -1 
	FOR EACH external_resource IN external_resources_array
		IF external_resource <> "" THEN array_position = array_position + 1
	NEXT
	
	'determining the size of the m-d array
	ReDim ext_res_multi_array(array_position, 5)
	
	'assigning values to the m-d array
	array_position = -1
	FOR EACH external_resource IN external_resources_array
		IF external_resource <> "" THEN 
			array_position = array_position + 1
			external_resource = split(external_resource, "|||")
			FOR i = 0 TO ubound(external_resource)
				ext_res_multi_array(array_position, i) = external_resource(i)			
			NEXT			
		END IF
	NEXT
	
	'dynamic dlg height
	dlg_height = 105 + (array_position * 15)
	
	'if the user has admin enabled, the dialog gets wider to accommodate 
	IF admin_enabled = TRUE THEN 
		dlg_width = 396
	ELSE
		dlg_width = 316
	END IF
	
	'building the dialog
	'this boss hogg is dynamic..it won't fit into dlgedit
    BeginDialog Dialog1, 0, 0, dlg_width, dlg_height, "External Resources"
        Text 5, 20, 55, 10, "Resource Name"
        Text 65, 20, 180, 10, "Description"
        
        dlg_row = 35
        button_number = 100
        FOR i = 0 TO array_position
        	ButtonGroup ButtonPressed
        	  PushButton 5, dlg_row, 55, 10, ext_res_multi_array(i, 0), button_number		'writing the button with the name on it
        	  ext_res_multi_array(i, 3) = button_number
        	Text 65, dlg_row, 225, 10, ext_res_multi_array(i, 2)							'writing the description
        	dlg_row = dlg_row + 15
        	button_number = button_number + 1
        NEXT
        ButtonGroup ButtonPressed
          CancelButton 5, dlg_height - 20, 50, 15
          
		'If this user is trusted with ADMIN privileges, they get access to these buttons. All other users will not have this access.
		IF admin_enabled = TRUE THEN 
			dlg_row = 35
			button_number = 1000
			'adding the UPDATE buttons
			FOR i = 0 TO array_position
				ButtonGroup ButtonPressed
				  PushButton 335, dlg_row, 15, 10, "O", button_number
				ext_res_multi_array(i, 4) = button_number
				button_number = button_number + 1
				dlg_row = dlg_row + 15
			NEXT
				
			dlg_row = 35
			button_number = 2000
			'adding the DELETE buttons
			FOR i = 0 TO array_position
				PushButton 365, dlg_row, 15, 10, "X", button_number
				ext_res_multi_array(i, 5) = button_number
				button_number = button_number + 1
				dlg_row = dlg_row + 15
			Next
			PushButton 330, 50 + (array_position * 15), 55, 15, "Add Resources", add_resource_button
			Text 330, 20, 25, 10, "Update"
			Text 360, 20, 25, 10, "Delete"
			GroupBox 320, 5, 70, 65 + (array_position * 15), "Admins ONLY"
		END IF
        EndDialog
		
		'running the dialog
		Dialog
			cancel_confirmation
			IF ButtonPressed = add_resource_button THEN CALL create_external_resource(default_directory)
			IF ButtonPressed >= 100 AND ButtonPressed < 1000 THEN 
				FOR i = 100 TO (100 + array_position)
					IF ButtonPressed = ext_res_multi_array(i - 100, 3) THEN 
						Set objExplorer = CreateObject("InternetExplorer.Application")
						objExplorer.Navigate ext_res_multi_array(i - 100, 1) 
						objExplorer.ToolBar = 1
						objExplorer.StatusBar = 1
						objExplorer.Visible = 1 
						
						EXIT FOR
					END IF
				NEXT
			END IF
			IF admin_enabled = TRUE THEN 
				'if the user presses the UPDATE button for that resource
				IF ButtonPressed >= 1000 AND ButtonPressed < 2000 THEN 
					FOR i = 1000 TO (1000 + array_position)
						IF ButtonPressed = ext_res_multi_array(i - 1000, 4) THEN						
							'the script needs to open the text file, find the link provided, and delete it
							With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
								If .FileExists(default_directory & "\External Resources.txt") Then										'If the file exists...
									Set TextFileObj = CreateObject("Scripting.FileSystemObject")										'Create another FSO
									Set text_command = TextFileObj.OpenTextFile(default_directory & "\External Resources.txt")			'Open the text file
									text_raw = text_command.ReadAll																		'Read the text file
									text_command.Close																					'Closes the file
								END IF
							END WITH

							'building the line_for_clearing
							original_line = ext_res_multi_array(i - 1000, 0) & "|||" & ext_res_multi_array(i - 1000, 1) & "|||" & ext_res_multi_array(i - 1000, 2)
							'calling function to update this resource
							CALL update_existing_resource(ext_res_multi_array(i - 1000, 0), ext_res_multi_array(i - 1000, 1), ext_res_multi_array(i - 1000, 2))   ', new_name, new_url, new_description)
							new_line = ext_res_multi_array(i - 1000, 0) & "|||" & ext_res_multi_array(i - 1000, 1) & "|||" & ext_res_multi_array(i - 1000, 2) 'new_name & "|||" & new_url & "|||" & new_description
							
							text_filtered = replace(text_raw, original_line, new_line)					'replacing the original text with the new text

							'Now it updates the text file
							With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
								Set TextFileObj = CreateObject("Scripting.FileSystemObject")											'Create another FSO
								Set text_command = TextFileObj.OpenTextFile(default_directory & "\External Resources.txt", 2)			'Open the text file for writing
								text_command.Write text_filtered			 															'Writes the new in a new file
								text_command.Close																						'Closes the file
							END WITH								
							
							EXIT FOR
						END IF				
					NEXT
				'if the user presses the DELETE button for that resource
				ELSEIF ButtonPressed >= 2000 THEN 
					FOR i = 2000 TO (2000 + array_position)
						IF ButtonPressed = ext_res_multi_array(i - 2000, 5) THEN 
							'the script needs to open the text file, find the link provided, and delete it
							With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
								If .FileExists(default_directory & "\External Resources.txt") Then										'If the file exists...
									Set TextFileObj = CreateObject("Scripting.FileSystemObject")										'Create another FSO
									Set text_command = TextFileObj.OpenTextFile(default_directory & "\External Resources.txt")			'Open the text file
									text_raw = text_command.ReadAll																		'Read the text file
									text_command.Close																					'Closes the file
								END IF
							END WITH

							'building the line_for_clearing
							line_for_clearing = ext_res_multi_array(i - 2000, 0) & "|||" & ext_res_multi_array(i - 2000, 1) & "|||" & ext_res_multi_array(i - 2000, 2)
							text_filtered = replace(text_raw, line_for_clearing, "")					'replacing the text with nothing

							'Now it updates the text file
							With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
								Set TextFileObj = CreateObject("Scripting.FileSystemObject")											'Create another FSO
								Set text_command = TextFileObj.OpenTextFile(default_directory & "\External Resources.txt", 2)			'Open the text file for writing
								text_command.Write text_filtered			 															'Writes the new in a new file
								text_command.Close																						'Closes the file
							END WITH								
							
							EXIT FOR
						END IF			
					NEXT
				END IF
			END IF
END FUNCTION

'this function is for updating an existing record
FUNCTION update_existing_resource(resource_name, resource_url, resource_description) ', new_name, new_url, new_description)
	    BeginDialog Dialog1, 0, 0, 346, 110, "Manage Resources"
	    	EditBox 95, 10, 110, 15, resource_name
	    	EditBox 95, 30, 200, 15, resource_url
	    	ButtonGroup ButtonPressed
	    	  PushButton 305, 35, 35, 10, "Test Link", test_link_button
	    	EditBox 95, 50, 200, 15, resource_description
	    	ButtonGroup ButtonPressed
	    	  OkButton 240, 90, 50, 15
	    	Text 10, 15, 65, 10, "Resource Name:"
	    	Text 10, 35, 65, 10, "Resource URL:"
	    	Text 10, 55, 75, 10, "Resource Description:"
	    EndDialog		
		DO
			err_msg = ""
			Dialog
				IF ButtonPressed = test_link_button THEN 
					Set objExplorer = CreateObject("InternetExplorer.Application")
					objExplorer.Navigate resource_url   
					objExplorer.ToolBar = 1
					objExplorer.StatusBar = 1
					objExplorer.Visible = 1             
				END IF
				'err_msg handling
				IF resource_name = "" THEN err_msg = err_msg & vbCr & "* You must provide a name for this resource. This is what will appear on the button."
				IF resource_url = "" THEN err_msg = err_msg & vbCr & "* You must provide a URL for this resource. This is where the resource will navigate."
				IF resource_description = "" AND ButtonPressed = -1 THEN msgbox "*** NOTICE!!! ***" & vbCr & vbCr & "You did not provide a description for this resource so your users know what to expect from this resource and why they would use it." & vbCr & vbCr & "This is not a required value and you may proceed without entering this value."
				IF err_msg <> "" THEN msgbox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1 AND err_msg = ""
END FUNCTION

'this function adds content to the txt file
FUNCTION create_external_resource(default_directory)
	DO
		'reseting the variables in the dialog
		resource_name = ""
		resource_url = ""
		resource_description = ""
		
		'building the dialog
	    BeginDialog Dialog1, 0, 0, 346, 110, "Manage Resources"
	    	EditBox 95, 10, 110, 15, resource_name
	    	EditBox 95, 30, 200, 15, resource_url
	    	ButtonGroup ButtonPressed
	    	  PushButton 305, 35, 35, 10, "Test Link", test_link_button
	    	EditBox 95, 50, 200, 15, resource_description
	    	ButtonGroup ButtonPressed
	    	  OkButton 240, 90, 50, 15
	    	  CancelButton 290, 90, 50, 15
	    	Text 10, 15, 65, 10, "Resource Name:"
	    	Text 10, 35, 65, 10, "Resource URL:"
	    	Text 10, 55, 75, 10, "Resource Description:"
	    EndDialog		
		
		DO
			err_msg = ""
			Dialog
				IF ButtonPressed = test_link_button THEN  			'opening the link 
					Set objExplorer = CreateObject("InternetExplorer.Application")
					objExplorer.Navigate resource_url   
					objExplorer.ToolBar = 1
					objExplorer.StatusBar = 1
					objExplorer.Visible = 1             
				END IF
				IF ButtonPressed = 0 THEN EXIT FUNCTION			'<<< exiting the function to give the admin user the option of going back to navigate around
				' err_msg handling...
				IF resource_name = "" THEN err_msg = err_msg & vbCr & "* You must provide a name for this resource. This is what will appear on the button."
				IF resource_url = "" THEN err_msg = err_msg & vbCr & "* You must provide a URL for this resource. This is where the resource will navigate."
				IF resource_description = "" AND ButtonPressed = -1 THEN msgbox "*** NOTICE!!! ***" & vbCr & vbCr & "You did not provide a description for this resource so your users know what to expect from this resource and why they would use it." & vbCr & vbCr & "This is not a required value and you may proceed without entering this value."
				IF err_msg <> "" THEN msgbox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1 AND err_msg = ""
		
		'declaring and opening the text file
		DIM oTxtFile
		WITH (CreateObject("Scripting.FileSystemObject"))
			If .FileExists(default_directory & "\External Resources.txt") Then			' <<< if the file exists then do the next gobbins
				Set external_resources_file = CreateObject("Scripting.FileSystemObject")
				Set external_resources_file_cmd = external_resources_file.OpenTextFile(default_directory & "\External Resources.txt")			'opening the text file
				external_resources_raw = external_resources_file_cmd.ReadAll																	'reading the text file
				external_resources_file_cmd.Close																								'closing the text file
				
				Set external_resources_file_command = external_resources_file.OpenTextFile(default_directory & "\External Resources.txt", 2)	're-opening the text file to write to it
			ELSE 
				' otherwise, if the file does not already exist, the script creates one
				Set external_resources_file = CreateObject("Scripting.FileSystemObject")
				Set external_resources_file_command = external_resources_file.CreateTextFile(default_directory & "\External Resources.txt", 2)	'creating the file
				external_resources_raw = ""
			END IF
				external_resources_file_command.Write external_resources_raw & vbNewLine & resource_name & "|||" & resource_url & "|||" & resource_description		'writing the file contents
				external_resources_file_command.Close																												'closing the file
		END WITH	
		do_it_again = MsgBox ("You have added " & resource_name & " to your agency's list of External Resources. Would you like to add another resource?", vbYesNo)	'seeing if the user wants to add another resource
	LOOP UNTIL do_it_again = vbNo
END FUNCTION


'=====================================
'===============The script============
'=====================================
'stopping the script if the agency does not have the default_directory value set...
IF default_directory = "" THEN script_end_procedure("Your agency's Global Variables file is not properly configured to use this script. You must provide a value for the variable ''default_directory'' for the script to run. Please contact a scripts administrator to resolve this issue.")

'grabbing the user's ID to determine if they should be able to update the links
Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = UCASE(objNet.UserName)

IF InStr(beta_users, windows_user_ID) <> 0 THEN 
	admin_enabled = TRUE
ELSE
	admin_enabled = FALSE
END IF

'the meat of the script
DO
	call external_resources_menu(default_directory, admin_enabled)				' <<<< All of the script is contained in this function. It needs to be in a function because of the dynamicalizing of the dialog and the reading and writing to the text file.
	do_it_again = MsgBox ("Do you need to access another resource? Press YES to continue. Press NO to stop the script.", vbYesNo)
LOOP WHILE do_it_again = vbYes

script_end_procedure("")

