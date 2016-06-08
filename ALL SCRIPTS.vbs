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
	
	public agencies_that_use	
end class

all_counties = "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

'The following is the list of all scripts. 
'>>>> THESE MUST BE MAINTAINED IN ALPHABETICAL ORDER ACCORDING TO THE SCRIPT_NAME PROPERTY <<<<<
script_num = 0 

ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ADJUSTMENTS"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - ADJUSTMENTS.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for documenting adjustments made to the case."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "AFFIDAVIT OF SERVICE DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - AFFIDAVIT OF SERVICE BY MAIL DOCS.vbs"
cs_scripts_array(script_num).description		= "Sends Affidavits of Service to multiple participants on the case."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ARREARS MGMT REVW"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - ARREARS MGMT REVIEW.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note for documenting an arrears management review."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CALI TO EXCEL"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - CALI TO EXCEL.vbs"
cs_scripts_array(script_num).description		= "Builds a list in Excel of case numbers, function types, program codes, interstate codes, and names on given CALI."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CASE TRANSFER"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CASE TRANSFER.vbs"
cs_scripts_array(script_num).description		= "Transfers single case and creates CAAD about why."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CLIENT CONTACT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CLIENT CONTACT.vbs"
cs_scripts_array(script_num).description		= "Creates a uniform CAAD note for when you have contact with or about client."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COLA"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - COLA.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!!! Leads you through performing a COLA. Adds CAAD note when completed."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COMPANION CASE FINDER - CP"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - CP COMPANION CASE FINDER.vbs"
cs_scripts_array(script_num).description		= "Builds list in Excel of companion cases for CPs on your CALI."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COMPANION CASE FINDER - NCP"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - NCP COMPANION CASE FINDER.vbs"
cs_scripts_array(script_num).description		= "Builds list in Excel of companion cases for NCPs on your CALI."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "COURT ORDER REQUEST"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - COURT ORDER REQUEST.vbs"
cs_scripts_array(script_num).description		= "Creates B0170 CAAD note for requesting a court order, which also creates worklist to remind worker of order request."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CP NAME CHANGE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - CP NAME CHANGE.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!! Updates CP name and alias. Modifies M1000 CAAD note."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "CSENET INFO"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - CSENET INFO.vbs"
cs_scripts_array(script_num).description		= "Creates T0111 CAAD note with text copied from INTD screen."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "DATE OF HEARING (EXPRO)"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - DATE OF THE HEARING (EXPRO).vbs"
cs_scripts_array(script_num).description		= "Date of hearing template for expro."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "DATE OF HEARING (JUDICIAL)"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - DATE OF THE HEARING (JUDICIAL).vbs"
cs_scripts_array(script_num).description		= "Date of hearing template for judicial."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "DDPL CALC"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - DDPL CALCULATOR.vbs"
cs_scripts_array(script_num).description		= "Calculates direct deposits made over user-provided date range."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "E-FILING"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - E-FILING.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!! Template for adding CAAD note about e-filing."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ENFORCEMENT INTAKE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ENFORCEMENT INTAKE.vbs"
cs_scripts_array(script_num).description		= "Intake workflow on enforcement cases."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "EST DORD NPA DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - EST DORD NPA DOCS.vbs"
cs_scripts_array(script_num).description		= "NEW 01/2016!! Generates DORD docs for NPA case."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ESTB DORD PA DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - ESTB DORD DOCS FOR PA CASE.vbs"
cs_scripts_array(script_num).description		= "NEW 01/2016!! Generates DORD docs for PA case."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FAILURE POF RSDI DFAS"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - FAILURE POF RSDI DFAS.vbs"
cs_scripts_array(script_num).description		= "Clears E0014 (Failure Notice to POF REVW) worklist when income is from RSDI or DFAS."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FEE SUPPRESSION OVERRIDE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - FEE SUPPRESSION OVERRIDE.vbs"
cs_scripts_array(script_num).description		= "Overrides a fee suppression."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FIND NAME ON CALI"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - FIND NAME ON CALI.vbs"
cs_scripts_array(script_num).description		= "Searches CALI for a specific CP or NCP."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "FRAUD REFERRAL"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - FRAUD REFERRAL.vbs"
cs_scripts_array(script_num).description		= "Template for adding CAAD note about a fraud referral."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "GENERIC ENFORCEMENT INTAKE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - GENERIC ENFORCEMENT INTAKE.vbs"
cs_scripts_array(script_num).description		= "Creates various docs related to CS intake as well as DORD docs and enters CAAD."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "INTAKE DOCS RECEIVED"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - INTAKE DOCS RECEIVED.vbs"
cs_scripts_array(script_num).description		= "NEW 02/2016!! Creates CAAD note for recording receipt of intake docs."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "IW CAAD CAWT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - IW CAAD CAWT.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!! Creates CAAD and CAWT about IW."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "IW CALC"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - IW CALCULATOR.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!! Calculator for determining the amount of IW over a given period."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "MES FINANCIAL DOCS SENT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - MES FINANCIAL DOCS SENT.vbs"
cs_scripts_array(script_num).description		= "NEW 02/2016!! Creates CAAD note for recording documents sent to parties."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NCP LOCATE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - NCP LOCATE.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!! Walks you through processing an NCP locate."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NON PAY"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - NON PAY.vbs"
cs_scripts_array(script_num).description		= "Sends DORD doc and creates CAAD related to Non-Pay."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NO PAY REPORT"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - NO PAY REPORT.vbs"
cs_scripts_array(script_num).description		= "Creates list in Excel of cases that have had no payment within given time period."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NO PAYMENT MONTHS ONE-FOUR"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - NO PAYMENT MONTHS ONE-FOUR.vbs"
cs_scripts_array(script_num).description		= "CAAD template documenting non-payment enforcement actions."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "NOTICE OF CONT'D SERVICE"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - EVALUATE NOCS.vbs"
cs_scripts_array(script_num).description		= "Evaluates D0800 (REVW for Notice of Cont'd Services) worklist and allows user to send DORD docs."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PALC CALC"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - PALC CALCULATOR.vbs"
cs_scripts_array(script_num).description		= "Calculates voluntary and involuntary amounts from the PALC screen."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PAYMENT PLAN REVIEW"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - PAYMENT PLAN REVIEW.vbs"
cs_scripts_array(script_num).description		= "CAAD template related to payment plan."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PAY OR REPORT"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - PAY OR REPORT.vbs"
cs_scripts_array(script_num).description		= "CAAD note for contempt/''pay or report'' instances."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "PRORATE SUPPORT"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - PRORATE SUPPORT.vbs"
cs_scripts_array(script_num).description		= "Calculator for determining pro-rated support for patrial months."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "QUARTERLY REVIEWS"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - QUARTERLY REVIEWS.vbs"
cs_scripts_array(script_num).description		= "CAAD note for quarterly review processes."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REDIRECT DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - DOCS FOR REDIRECT.vbs"
cs_scripts_array(script_num).description		= "Creates redirection docs and redirection worklist items."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REVW PAY PLAN - DL IS SUSP"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - E4111 SUSP SCRUBBER.vbs"
cs_scripts_array(script_num).description		= "NEW 02/2016!! Scrubs E4111 (REVW Pay Plan) workflists when DL is already suspended."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REVW PAY PLAN RECENT ACTIVITY"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - E4111 WORKLIST SCRUBBER.vbs"
cs_scripts_array(script_num).description		= "NEW 02/2016!! Presents recent payment activity to evaluate E4111 (REVW Pay Plan) worklists."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "REVW QW INFO"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - REVIEW QW INFO.vbs"
cs_scripts_array(script_num).description		= "NEW 01/2016!! Purgres all M8001 (REVW Case Referred) worklist from your USWT."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "ROP DETAIL"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - ROP DETAIL.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note about the date parties signed Recognition of Parentage."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "SOP INVOICE"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - SOP INVOICE.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note that the Service of Process invoice was received, details about the service, and if the invoice is OK to pay."
cs_scripts_array(script_num).agencies_that_use		= "BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "TRANSFER CASE(S)"
cs_scripts_array(script_num).script_type		= "BULK"
cs_scripts_array(script_num).file_name			= "BULK - CASE TRANSFER.vbs"
cs_scripts_array(script_num).description		= "Script for transfering multiple cases, from one caseload to multiple caseloads."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "UN UN DOCS"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - UNREIMBURSED UNINSURED DOCS.vbs"
cs_scripts_array(script_num).description		= "Prints DORD docs for collecting unreimbursed and uninsured expenses."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "UPDATE WORKER SIGNATURE"
cs_scripts_array(script_num).script_type		= "ACTIONS"
cs_scripts_array(script_num).file_name			= "ACTIONS - UPDATE WORKER SIGNATURE.vbs"
cs_scripts_array(script_num).description		= "NEW 04/2016!! Allows you to maintain a default signature that loads in all scripts."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

script_num = script_num + 1
ReDim Preserve cs_scripts_array(script_num)
SET cs_scripts_array(script_num) = NEW cs_script
cs_scripts_array(script_num).script_name		= "WAIVER OF PERSONAL SERVICE"
cs_scripts_array(script_num).script_type		= "NOTES"
cs_scripts_array(script_num).file_name			= "NOTES - WAIVER OF PERSONAL SERVICE.vbs"
cs_scripts_array(script_num).description		= "Creates CAAD note of the date a CP signed the waiver of personal service document."
cs_scripts_array(script_num).agencies_that_use		= "ANOKA, BELTRAMI, DAKOTA, HENNEPIN, MILLE LACS, OLMSTED, RAMSEY, RENVILLE, SCOTT, STEARNS, WASHINGTON, WRIGHT"

