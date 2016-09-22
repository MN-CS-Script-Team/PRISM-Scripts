'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------
'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the installer, but doing so should not be undertaken lightly.

' DETAILS ABOUT HOW YOUR SCRIPTS WILL RUN -------'---------------------------------------------------------------------------

'Run locally: if this is set to "True", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts. Only scriptwriters should do it. An agency should always be set to "false".
run_locally = true

'This is a variable which signifies the agency uses the master branch or the RELEASE branch. Set to true if you're a scriptwriter agency and all users are going to be on the master branch. Otherwise, set to false.
use_master_branch = true

'This allows a "beta user" group to have access to master branch scripts, while everyone else uses release. This is helpful for counties that want to maintain a small test group.
'Here is the list of agency super users. These users will have access to the test scripts. Enter the list of users' log-in IDs in the quotes below, comma separated
beta_users = ""

'This is used by the AGENCY CUSTOMIZED process, and can be used elsewhere if needed, but for now it's mostly informational
'	This is modified by the installer, which will determine if this is a scriptwriter or a production user.
default_directory = "C:\DHS-PRISM-Scripts\Script Files\"

'DETAILS ABOUT STATISTICS AND GATHERING THEM ------------------------------------------------------------------------------------------

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = False

'This is the file path for the statistics Access database.
stats_database_path = "C:\DHS-PRISM-Scripts\Databases for script usage\usage statistics.accdb"

'DETAILS ABOUT WHERE TO FIND DOCS AND WHICH TO USE ------------------------------------------------------------------------------------------

'This is the folder path for county-specific Word documents. Modify this with your shared-drive location for Word documents.
word_documents_folder_path = "C:\DHS-PRISM-Scripts\Word files for script usage\"

'DETAILS ABOUT THE COUNTY ITSELF -------------------------------------------------------------------------------------------------------------

'This is the county code on the CALI screen.
county_cali_code = "###"

'An array of county attorneys. "Select one:" should ALWAYS be in there, and ALWAYS be first. Replace "County Attorney #" with your agency's county attorney names.
county_attorney_array = array("County Attorney 1", "County Attorney 2", "County Attorney 3", "County Attorney 4", "County Attorney 5")

'An array of child support magistrates. "Select one:" should ALWAYS be in there, and ALWAYS be first.  Replace "Magistrate # with your agency's child support magistrate names.
child_support_magistrates_array = array("Magistrate 1", "Magistrate 2", "Magistrate 3", "Magistrate 4", "Magistrate 5")

'An array of judges. "Select one:" should ALWAYS be in there, and ALWAYS be first.  Replace "Judge #" with your agency's judges names.
county_judge_array = array("Judge 1", "Judge 2", "Judge 3", "Judge 4", "Judge 5")

'This is used by scripts which tell the worker where to find a doc to send to a client (ie "Send form using Compass Pilot")
EDMS_choice = "Compass Pilot"

'This is the county's email support address. It can be a distribution list or an individual.
support_email_address = "jean.valjean@paris.fr"

'ACTIONS TAKEN BASED ON COUNTY CUSTOM VARIABLES------------------------------------------------------------------------------
'**DO NOT EDIT BELOW THIS LINE UNLESS YOU ARE ABSOLUTELY SURE OF WHAT YOU ARE DOING**

is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'This loads the user ID for use in determining beta users. May also be used elsewhere in scripts.
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName

'This will assign beta users to the master branch.
If InStr(UCASE(beta_users), UCASE(windows_user_ID)) <> 0 then use_master_branch = true

'This is the URL of our script repository, and should only change if the agency is a scriptwriting agency. Scriptwriters can elect to use the master branch, allowing them to test new tools, etc.
IF use_master_branch = TRUE THEN		'scriptwriters typically use the master branch
	script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/Script Files/"
ELSE							'Everyone else (who isn't a scriptwriter) typically uses the release branch
	script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/release/Script Files/"
END IF

'If run locally is set to "True", the scripts will totally bypass GitHub and run locally.
IF run_locally = TRUE THEN script_repository = "C:\DHS-PRISM-Scripts\Script Files\"
