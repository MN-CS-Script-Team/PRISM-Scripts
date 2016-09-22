'A class for each script item. This will need to be moved to a global files location.
class cs_script
	public script_name
	public file_name
	public description
	public button
	public script_type

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 4.3 ) + 10
	end property

end class

'The following is the list of all scripts.
'>>>> THESE MUST BE MAINTAINED IN ALPHABETICAL ORDER ACCORDING TO THE SCRIPT_NAME PROPERTY <<<<<
script_num = 0

ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ADJUSTMENTS"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - ADJUSTMENTS.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for documenting adjustments made to the case."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ADMIN REDIRECT"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ADMIN REDIRECT.vbs"
cs_scripts_array(script_num).description		= "Creates redirection docs and redirection worklist items."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "AFFIDAVIT OF SERVICE BY MAIL DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - AFFIDAVIT OF SERVICE BY MAIL DOCS.vbs"
cs_scripts_array(script_num).description		= "Sends Affidavits of Service to multiple participants on the case."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ANOKA SANCTION"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ANOKA SANCTION.vbs"
cs_scripts_array(script_num).description		= "Takes actions on the case to apply or remove public assistance sanction for non-cooperation with child support."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ARREARS MGMT REVIEW"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - ARREARS MGMT REVIEW.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for documenting an arrears management review."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CAAD"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - CAAD.vbs"
cs_scripts_array(script_num).description		= "Navigates to the CAAD screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CAFS"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - CAFS.vbs"
cs_scripts_array(script_num).description		= "Navigates to the CAFS screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CALI TO EXCEL"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - CALI TO EXCEL.vbs"
cs_scripts_array(script_num).description		= "Builds a list in Excel of case numbers, function types, program codes, interstate codes, and names on given CALI."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CAPS"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - CAPS.vbs"
cs_scripts_array(script_num).description		= "Navigates to the CAPS screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CASE INITIATION DOCS RECEIVED"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CASE INITIATION DOCS RECEIVED.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording receipt of intake docs."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CASE TRANSFER"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - CASE TRANSFER.vbs"
cs_scripts_array(script_num).description		= "Transfers single case and creates CAAD about why."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CASE TRANSFER"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - CASE TRANSFER.vbs"
cs_scripts_array(script_num).description		= "Gives the user the ability to quickly transfer mulitple cases."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CAST"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - CAST.vbs"
cs_scripts_array(script_num).description		= "Navigates to the CAST screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CAWT"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - CAWT.vbs"
cs_scripts_array(script_num).description		= "Navigates to the CAWT screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CLIENT CONTACT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CLIENT CONTACT.vbs"
cs_scripts_array(script_num).description		= "Creates a uniform CAAD note for when you have contact with or about client."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COLA"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - COLA.vbs"
cs_scripts_array(script_num).description		= "Leads you through performing a COLA. Adds CAAD note when completed."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CONTEMPT HEARING"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CONTEMPT HEARING.vbs"
cs_scripts_array(script_num).description		= "Creates a hearing date CAAD note for a contempt hearing."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COURT ORDER REQUEST"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - COURT ORDER REQUEST.vbs"
cs_scripts_array(script_num).description		= "Creates B0170 CAAD note for requesting a court order, which also creates worklist to remind worker of order request."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COURT PREP WORKSHEET"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - COURT PREP WORKSHEET.vbs"
cs_scripts_array(script_num).description		= "Runs a court prep worksheet in anticipation of court dates, getting info from PRISM and putting it into a Word doc."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CP COMPANION CASE FINDER"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - CP COMPANION CASE FINDER.vbs"
cs_scripts_array(script_num).description		= "Builds list in Excel of companion cases for CPs on your CALI."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CPDD"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - CPDD.vbs"
cs_scripts_array(script_num).description		= "Navigates to the CPDD screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CSENET INFO"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CSENET INFO.vbs"
cs_scripts_array(script_num).description		= "Creates T0111 CAAD note with text copied from INTD screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "DDPL CALCULATOR"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - DDPL CALCULATOR.vbs"
cs_scripts_array(script_num).description		= "Calculates direct deposits made over user-provided date range."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "E-FILING"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - E-FILING.vbs"
cs_scripts_array(script_num).description		= "Template for adding CAAD note about e-filing."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "EMC DORD DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - EMC DORD DOCS.vbs"
cs_scripts_array(script_num).description		= "Sends emancipation DORD docs."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "EMPLOYMENT VERIFICATION"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - EMPLOYMENT VERIFICATION.vbs"
cs_scripts_array(script_num).description		= "NEW 08/2016!! - Complete an Employment Verification in NCID or CPID, includes info on CAAD note."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ENFL"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - ENFL.vbs"
cs_scripts_array(script_num).description		= "Navigates to the ENFL screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ENFORCEMENT INTAKE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ENFORCEMENT INTAKE.vbs"
cs_scripts_array(script_num).description		= "Intake workflow on enforcement cases."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ESTB DORD DOCS FOR NPA CASE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ESTB DORD DOCS FOR NPA CASE.vbs"
cs_scripts_array(script_num).description		= "Generates establishment DORD docs for NPA case."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ESTB DORD DOCS FOR PA CASE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ESTB DORD DOCS FOR PA CASE.vbs"
cs_scripts_array(script_num).description		= "Generates establishment DORD docs for PA case."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FAILURE POF RSDI DFAS"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - FAILURE POF RSDI DFAS.vbs"
cs_scripts_array(script_num).description		= "Clears E0014 (Failure Notice to POF REVW) worklist when income is from RSDI or DFAS."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FEE SUPPRESSION OVERRIDE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - FEE SUPPRESSION OVERRIDE.vbs"
cs_scripts_array(script_num).description		= "Overrides a fee suppression."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FIND NAME ON CALI"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - FIND NAME ON CALI.vbs"
cs_scripts_array(script_num).description		= "Searches CALI for a specific CP or NCP."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FIND THAT SCREEN IN PRISM"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - FIND THAT SCREEN IN PRISM.vbs"
cs_scripts_array(script_num).description		= "Displays a list of PRISM screens which you can then select."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FRAUD REFERRAL"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - FRAUD REFERRAL.vbs"
cs_scripts_array(script_num).description		= "Template for adding CAAD note about a fraud referral."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "GENERIC ENFORCEMENT INTAKE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - GENERIC ENFORCEMENT INTAKE.vbs"
cs_scripts_array(script_num).description		= "Creates various docs related to CS intake as well as DORD docs and enters CAAD."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "HEARING NOTES"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - HEARING NOTES.vbs"
cs_scripts_array(script_num).description		= "NEW 08/2016!! - CAAD note template for sending details about hearing notes."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "INFO"
cs_scripts_array(script_num).script_type		= "UTILITIES"
cs_scripts_array(script_num).file_name			= "UTILITIES - INFO.vbs"
cs_scripts_array(script_num).description		= "Displays information about your BlueZone Scripts installation."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "INVOICES"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - INVOICES.vbs"
cs_scripts_array(script_num).description		= "NEW 07/2016!! - Creates CAAD note for recording invoices."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "IW CAAD CAWT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - IW CAAD CAWT.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD and CAWT about IW."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "IW CALCULATOR"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - IW CALCULATOR.vbs"
cs_scripts_array(script_num).description		= "Calculator for determining the amount of IW over a given period."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "MAXIS SCREEN FINDER"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - MAXIS SCREEN FINDER.vbs"
cs_scripts_array(script_num).description		= "Displays a list of MAXIS screens you can select."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "MES FINANCIAL DOCS SENT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - MES FINANCIAL DOCS SENT.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording documents sent to parties."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "MOD CAAD NOTE - CONTACT CHECKLIST"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - MOD CAAD NOTE - CONTACT CHECKLIST.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for recording contact with Client regarding possible Mod."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NCDD"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - NCDD.vbs"
cs_scripts_array(script_num).description		= "Navigates to the NCDD screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NCID"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - NCID.vbs"
cs_scripts_array(script_num).description		= "Navigates to the NCID screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NCP COMPANION CASE FINDER"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - NCP COMPANION CASE FINDER.vbs"
cs_scripts_array(script_num).description		= "Builds list in Excel of companion cases for NCPs on your CALI."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NCP LOCATE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - NCP LOCATE.vbs"
cs_scripts_array(script_num).description		= "Walks you through processing an NCP locate."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NO PAY REPORT"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - NO PAY REPORT.vbs"
cs_scripts_array(script_num).description		= "Creates list in Excel of cases that have had no payment within given time period."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NON PAY"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - NON PAY.vbs"
cs_scripts_array(script_num).description		= "Sends DORD doc and creates CAAD related to Non-Pay."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NOTICE OF CONTINUED SERVICE"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - NOTICE OF CONTINUED SERVICE.vbs"
cs_scripts_array(script_num).description		= "Evaluates D0800 (REVW for Notice of Cont'd Services) worklist and allows user to send DORD docs."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PALC"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - PALC.vbs"
cs_scripts_array(script_num).description		= "Navigates to the PALC screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PALC CALCULATOR"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - PALC CALCULATOR.vbs"
cs_scripts_array(script_num).description		= "Calculates voluntary and involuntary amounts from the PALC screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PAPL"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - PAPL.vbs"
cs_scripts_array(script_num).description		= "Navigates to the PAPL screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PAY OR REPORT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - PAY OR REPORT.vbs"
cs_scripts_array(script_num).description		= "CAAD note for contempt/''pay or report'' instances."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PESE"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - PESE.vbs"
cs_scripts_array(script_num).description		= "Navigates to the PESE screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PRORATE SUPPORT"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - PRORATE SUPPORT.vbs"
cs_scripts_array(script_num).description		= "Calculator for determining pro-rated support for patrial months."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "QUARTERLY REVIEWS"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - QUARTERLY REVIEWS.vbs"
cs_scripts_array(script_num).description		= "CAAD note for quarterly review processes."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "QUICK CAAD"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - QUICK CAAD.vbs"
cs_scripts_array(script_num).description		= "NEW 08/2016!! - Quickly add links to CAAD codes you frequently use. Includes a search feature."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "RECORD IW INFO"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - RECORD IW INFO.vbs"
cs_scripts_array(script_num).description		= "NEW 09/2016!! - Record IW withholding info in a CAAD note, worklist, or view in a message box."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "RETURNED MAIL"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - RETURNED MAIL.vbs"
cs_scripts_array(script_num).description		= "NEW 09/2016!! - Updates address to new or unknown, and creates CAAD note."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REVIEW QW INFO"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - REVIEW QW INFO.vbs"
cs_scripts_array(script_num).description		= "Reviews all L2500 and L2501 worklists on your caseload and purges the worklist if the employer is already on NCID/CPID."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REVW PAY PLAN - DL IS SUSP"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - REVW PAY PLAN - DL IS SUSP.vbs"
cs_scripts_array(script_num).description		= "Scrubs E4111 (REVW Pay Plan) workflists when DL is already suspended."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REVW PAY PLAN RECENT ACTIVITY"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - REVW PAY PLAN RECENT ACTIVITY.vbs"
cs_scripts_array(script_num).description		= "Presents recent payment activity to evaluate E4111 (REVW Pay Plan) worklists."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "SEND F0104 DORD MEMO"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - SEND F0104 DORD MEMO.vbs"
cs_scripts_array(script_num).description		= "Sends F0104 DORD Memo Docs, with options to send a memo to both parties and preview memo text."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "SUCW"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - SUCW.vbs"
cs_scripts_array(script_num).description		= "Navigates to the SUCW screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "UNREIMBURSED UNINSURED DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - UNREIMBURSED UNINSURED DOCS.vbs"
cs_scripts_array(script_num).description		= "Prints DORD docs for collecting unreimbursed and uninsured expenses."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "UPDATE WORKER SIGNATURE"
cs_scripts_array(script_num).script_type		= "UTILITIES"
cs_scripts_array(script_num).file_name			= "UTILITIES - UPDATE WORKER SIGNATURE.vbs"
cs_scripts_array(script_num).description		= "Allows you to maintain a default signature that loads in all scripts."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "USWD"
cs_scripts_array(script_num).script_type		= "NAV"
cs_scripts_array(script_num).file_name			= "NAV - USWD.vbs"
cs_scripts_array(script_num).description		= "Navigates to the USWD screen."

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "WAIVER OF PERSONAL SERVICE"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - WAIVER OF PERSONAL SERVICE.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note of the date a CP signed the waiver of personal service document."
