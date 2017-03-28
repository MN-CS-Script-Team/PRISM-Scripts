'A class for each script item. This will need to be moved to a global files location.
class cs_script
	public script_name
	public description
	public button
	public category
	public release_date				'This allows the user to indicate when the script goes live (controls NEW!!! messaging)
	public scriptwriter				'Simply informational as of 10/2016.

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 4.3 ) + 10
	end property

	public property get file_name
		file_name = lcase(replace(script_name, " ", "-")) & ".vbs"
	end property
end class


'This variable must always start at zero. It figures out how many buttons it needs to process and create.
script_num = 0

'The following is the list of all scripts.

'ACTIONS SCRIPTS ========================================================================================================================================================================='

ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Admin Redirect"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Creates redirection docs and redirection worklist items."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Affidavit of Service by Mail Docs"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Sends Affidavits of Service to multiple participants on the case."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Case Transfer"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Transfers single case and creates CAAD about why."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COLA"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Leads you through performing a COLA. Adds CAAD note when completed."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Emancipation DORD docs"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Sends emancipation DORD docs."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Employment Verification"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Complete an Employment Verification in NCID or CPID, includes info on CAAD note."
cs_scripts_array(script_num).release_date		= #08/01/2016#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Enforcement Intake"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Intake workflow on enforcement cases."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Establishment DORD docs - NPA"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Generates establishment DORD docs for NPA case."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Establishment DORD docs - PA"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Generates establishment DORD docs for PA case."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Fee Suppression Override"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Overrides a fee suppression."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script				
cs_scripts_array(script_num).script_name		= "Financial Statement Follow-up"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Sends follow-up memo to parties regarding financial statements."
cs_scripts_array(script_num).release_date		= #11/14/2016#
cs_scripts_array(script_num).scriptwriter		= ""													
													
script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Find Name on CALI"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Searches CALI for a specific CP or NCP."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Generic Enforcement Intake"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Creates various docs related to CS intake as well as DORD docs and enters CAAD."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""
														
script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script				
cs_scripts_array(script_num).script_name		= "Income Verification"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Generates Word document regarding payments CP has received on their case."
cs_scripts_array(script_num).release_date		= #11/14/2016#
cs_scripts_array(script_num).scriptwriter		= ""																

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Interview Information Sheet"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Creates a Word document with general and case-specific information to be used as a reference when meeting with clients."
cs_scripts_array(script_num).release_date		= #01/31/2017#
cs_scripts_array(script_num).scriptwriter		= ""
																	
script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script				
cs_scripts_array(script_num).script_name		= "Interview Information Sheet"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Generates Word document of high-level case information and county-specific resources."
cs_scripts_array(script_num).release_date		= #11/23/2016#
cs_scripts_array(script_num).scriptwriter		= ""	

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NCP Locate"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Walks you through processing an NCP locate."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Non Pay"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Sends DORD doc and creates CAAD related to Non-Pay."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Record IW Info"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Record IW withholding info in a CAAD note, worklist, or view in a message box."
cs_scripts_array(script_num).release_date		= #09/01/2016#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Returned Mail"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Updates address to new or unknown, and creates CAAD note."
cs_scripts_array(script_num).release_date		= #09/01/2016#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Sanction"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Takes actions on the case to apply or remove public assistance sanction for non-cooperation with child support."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Send F0104 DORD memo"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Sends F0104 DORD Memo Docs, with options to send a memo to both parties and preview memo text."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Unreimbursed Uninsured Returned Docs"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Sends DORD docs when unreimbursed and uninsured docs are returned."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Unreimbursed Uninsured Sending Docs"
cs_scripts_array(script_num).category			= "actions"
cs_scripts_array(script_num).description		= "Prints DORD docs for collecting unreimbursed and uninsured expenses."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

'BULK SCRIPTS ============================================================================================================================================================================'

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Case Transfer"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Gives the user the ability to quickly transfer mulitple cases."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Failure POF RSDI DFAS"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Clears E0014 (Failure Notice to POF REVW) worklist when income is from RSDI or DFAS."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "List Generator - CALI"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Builds a list in Excel of case numbers, function types, program codes, interstate codes, and names on given CALI."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "List Generator - Companion Cases - CP"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Builds list in Excel of companion cases for CPs on your CALI."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "List Generator - Companion Cases - NCP"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Builds list in Excel of companion cases for NCPs on your CALI."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "List Generator - No Pay"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Creates list in Excel of cases that have had no payment within given time period."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Notice of Continued Service"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Evaluates D0800 (REVW for Notice of Cont'd Services) worklist and allows user to send DORD docs."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script				
cs_scripts_array(script_num).script_name		= "PA Program Reopen-Review"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Reviews M1600 worklist and purges cases that were closed for a reason in which we would not reopen."
cs_scripts_array(script_num).release_date		= #11/14/2016#
cs_scripts_array(script_num).scriptwriter		= ""		

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Review Pay Plan - DL is Suspended"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Scrubs E4111 (REVW Pay Plan) workflists when DL is already suspended."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Review Pay Plan - Recent Activity"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Presents recent payment activity to evaluate E4111 (REVW Pay Plan) worklists."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Review QW Info"
cs_scripts_array(script_num).category			= "bulk"
cs_scripts_array(script_num).description		= "Reviews all L2500 and L2501 worklists on your caseload and purges the worklist if the employer is already on NCID/CPID."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""



'CALCULATOR SCRIPTS ======================================================================================================================================================================='

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "DDPL"
cs_scripts_array(script_num).category			= "calculators"
cs_scripts_array(script_num).description		= "Calculates direct deposits made over user-provided date range."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "IW"
cs_scripts_array(script_num).category			= "calculators"
cs_scripts_array(script_num).description		= "Calculator for determining the amount of IW over a given period."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PALC"
cs_scripts_array(script_num).category			= "calculators"
cs_scripts_array(script_num).description		= "Calculates voluntary and involuntary amounts from the PALC screen."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Prorate Support"
cs_scripts_array(script_num).category			= "calculators"
cs_scripts_array(script_num).description		= "Calculator for determining pro-rated support for patrial months."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""




'NOTES SCRIPTS ======================================================================================================================================================================='


script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Adjustments"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for documenting adjustments made to the case."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Arrears Management Review"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for documenting an arrears management review."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Case Initiation Docs Received"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording receipt of intake/case initiation docs."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Client Contact"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates a uniform CAAD note for when you have contact with or about client."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Contempt Hearing"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates a hearing date CAAD note for a contempt hearing."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Court Order Request"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates B0170 CAAD note for requesting a court order, which also creates worklist to remind worker of order request."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CSENET Info"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates T0111 CAAD note with text copied from INTD screen."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "E-Filing"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Template for adding CAAD note about e-filing."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Fraud Referral"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Template for adding CAAD note about a fraud referral."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Hearing Notes"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "CAAD note template for sending details about hearing notes."
cs_scripts_array(script_num).release_date		= #08/01/2016#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Invoices"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording invoices."
cs_scripts_array(script_num).release_date		= #07/01/2016#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "IW CAAD CAWT"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD and CAWT about IW."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Maintaining County"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for requesting maintaining county."
cs_scripts_array(script_num).release_date		= #02/22/2017#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "MES Financial Docs Sent"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording documents sent to parties."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Mod CAAD Note - Contact Checklist"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording contact with Client regarding possible Mod."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Pay or Report"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "CAAD note for contempt/''pay or report'' instances."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Quarterly Reviews"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "CAAD note for quarterly review processes."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Waiver of Personal Service"
cs_scripts_array(script_num).category			= "notes"
cs_scripts_array(script_num).description		= "Creates CAAD note of the date a CP signed the waiver of personal service document."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""







'UTILITIES SCRIPTS ======================================================================================================================================================================='

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Changelog"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "A script which generates a changelog for end users."
cs_scripts_array(script_num).release_date		= #11/01/2016#
cs_scripts_array(script_num).scriptwriter		= "Veronica Cary and Robert Fewins-Kalb"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Email Scripts Support"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "Sends an email to your designated support person (specified by the installer)."
cs_scripts_array(script_num).release_date		= #10/01/2016#
cs_scripts_array(script_num).scriptwriter		= "Veronica Cary"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "External Resources"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "An agency-customizable list of web resources for general use."
cs_scripts_array(script_num).release_date		= #11/01/2016#
cs_scripts_array(script_num).scriptwriter		= "Robert Fewins-Kalb"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "MAXIS Screen Finder"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "Displays a list of MAXIS screens you can select."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PRISM Screen Finder"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "Displays a list of PRISM screens which you can then select."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Quick CAAD"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "Quickly add links to CAAD codes you frequently use. Includes a search feature."
cs_scripts_array(script_num).release_date		= #08/01/2016#
cs_scripts_array(script_num).scriptwriter		= ""

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Scripts Install Info"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "Displays information about your BlueZone Scripts installation."
cs_scripts_array(script_num).release_date		= #08/01/2016#
cs_scripts_array(script_num).scriptwriter		= "Veronica Cary"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "Update Worker Signature"
cs_scripts_array(script_num).category			= "utilities"
cs_scripts_array(script_num).description		= "Allows you to maintain a default signature that loads in all scripts."
cs_scripts_array(script_num).release_date		= #01/01/2000#
cs_scripts_array(script_num).scriptwriter		= "Robert Fewins-Kalb"
