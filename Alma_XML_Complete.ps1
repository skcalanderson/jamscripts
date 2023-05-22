# 2016-10-19
# Al Anderson
#
# This PowerShell script is designed to demonstrate how to run a query
# against a database and dump to a csv
#
# Usage:  Save this file as C:\sandbox\powershell\databaseQuery.ps1
# Start Powershell (Win-R, powershell)
# Execute the script (C:\sandbox\powershell\databaseQuery.ps1)

#
# If the above fails due to
# "cannot be loaded because the execution of scripts is disabled on this system"
# run this command within Powershell
# Set-ExecutionPolicy RemoteSigned
# See also http://technet.microsoft.com/en-us/library/ee176949.aspx
#
# 12-20-16 ala
# added case statement to select statement to change group code to what Alma wants
# added case statement to select statement to change phone to 0 if there is a NULL number
# added if statement around phone XML element to not include if phone is 0
# change id_type to 'INST_ID'

# http://www.vistax64.com/powershell/190352-executing-sql-queries-powershell.html
$server = "x.x.x.x"
$database = "tmseprd"
#$query = "SELECT id_num,x_first,x_last,email_address,x_gender,x_mobile_phone,x_7to16,x_landline_phone1,x_18,x_landline_phone2 FROM skc_rave_extract_emplye_stdnt_v"
$query = @"
  SELECT DISTINCT name_master.id_num,   
         name_master.first_name,   
         name_master.last_name,
		 NAME_MASTER.PREFERRED_NAME,
		 NAME_MASTER.MIDDLE_NAME,   
         name_master.email_address,
		 ADDRESS_MASTER.ADDR_LINE_1,
		 ADDRESS_MASTER.CITY,   
		 ADDRESS_MASTER.STATE,
		 ADDRESS_MASTER.ZIP,
		 CASE
			WHEN ADDRESS_MASTER.PHONE is null THEN 0
			ELSE ADDRESS_MASTER.PHONE
		 END as PHONE,        
		 	CASE 
			    WHEN EMPL_MAST.GRP_CDE = 'STAFF' THEN 'pstaff' 
			    ELSE 'pfaculty'
		 END as GRP_CDE,
		 TW_WEB_SECURITY.WEB_LOGIN
    FROM address_master LEFT OUTER JOIN empl_wrk_loc ON address_master.id_num = empl_wrk_loc.id_num,   
         empl_mast,   
         ind_pos_hist,   
         name_master,   
         biograph_master,
		 TW_WEB_SECURITY
   WHERE ( ind_pos_hist.id_num = empl_mast.id_num ) and  
		 ( tw_web_security.ID_NUM = empl_mast.ID_NUM) and
         ( empl_mast.id_num = name_master.id_num ) and  
         ( name_master.id_num = address_master.id_num ) and  
         ( name_master.current_address = address_master.addr_cde ) and  
         ( empl_mast.id_num = biograph_master.id_num ) and  
         ( ( empl_mast.act_inact_sts <> 'I' ) AND  
         ( ind_pos_hist.pos_sts <> 'I' ) AND  
         (ind_pos_hist.pos_end_dte is NULL OR  
         ind_pos_hist.pos_end_dte >= getdate() ) AND  
         (empl_mast.grp_cde in ('ADJUN','STAFF','FCLTY')) AND  
         (ind_pos_hist.org_pos not in ('CONHR','CONTR')) AND  
         name_master.email_address like '%@skc.edu' ) 

UNION ALL

SELECT DISTINCT name_master.id_num,   
         name_master.first_name,
         name_master.last_name,
		 NAME_MASTER.PREFERRED_NAME,
		 NAME_MASTER.MIDDLE_NAME,   
         name_master.email_address,
		 ADDRESS_MASTER.ADDR_LINE_1,
		 ADDRESS_MASTER.CITY,   
		 ADDRESS_MASTER.STATE,
		 ADDRESS_MASTER.ZIP,
		 CASE
			WHEN ADDRESS_MASTER.PHONE is null THEN 0
			ELSE ADDRESS_MASTER.PHONE
		 END as PHONE,      
		 'pstudent' as GRP_CDE,
		 TW_WEB_SECURITY.WEB_LOGIN
    FROM stud_term_sum_div,   
         name_master,   
         student_master,   
         biograph_master,   
         address_master,
         reg_config,  
		 TW_WEB_SECURITY
   WHERE ( biograph_master.id_num = name_master.id_num ) and  
		 ( tw_web_security.ID_NUM = biograph_master.ID_NUM) and
         ( name_master.id_num = address_master.id_num ) and  
         ( name_master.current_address = address_master.addr_cde ) and  
         ( name_master.id_num = student_master.id_num ) and  
         ( student_master.id_num = stud_term_sum_div.id_num ) and  
         ( ( stud_term_sum_div.div_cde in ('GR', 'UG') ) AND  
         ( stud_term_sum_div.yr_cde = reg_config.cur_yr_dflt ) AND  
         ( stud_term_sum_div.trm_cde = reg_config.cur_trm_dflt ) AND  
         ( student_master.loc_cde in ('GE','GN','PA' )) AND  
         ( stud_term_sum_div.pt_ft_hrs > 0 ) AND  
         ( name_master.email_address is not NULL ) AND  
         ( name_master.email_address like '%student.skc.edu' ) )   


"@
		 
		 
# We need to create the expiry date and purge date based on the current date.
$dte = Get-Date
$currentYear = $dte.Year
if ($dte.Month -gt 6)
{
    $expdte = Get-Date -Year ($currentYear + 1) -Month 6 -Day 30
} else
{
    $expdte = Get-Date -Year $currentYear -Month 6 -Day 30
}
# This is our expiry data
$expdteStr = $expdte.ToString("yyyy-MM-ddZ")
$purgeDte = $expdte.AddYears(1)
$purgeDteStr = $purgeDte.ToString("yyyy-MM-ddZ")



# powershell raw/verbatim strings are ugly
# Update this with the actual path where you want data dumped
#$extractFile = @"C:\Users\Al Anderson\dev\rave-people.csv
#"@

# If you have to use users and passwords, my condolences
$connectionTemplate = "Data Source={0};User ID=******;Password=********;Initial Catalog={1};"
$connectionString = [string]::Format($connectionTemplate, $server, $database)
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

$command = New-Object System.Data.SqlClient.SqlCommand
$command.CommandText = $query
$command.Connection = $connection

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $command
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$connection.Close()

# foreach ($row in $DataSet.Tables[0].Rows)
# {
	# Write-Host $row.id_num
# }

$XMLPath = "C:\Alma-Upload-Scripts\skc_alma.xml"

# get an XMLTextWriter to create the XML
$XmlWriter = New-Object System.XMl.XmlTextWriter($XMLPath,$Null)
 
# choose a pretty formatting:
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$xmlWriter.IndentChar = "`t"
# 
 
write the header
$xmlWriter.WriteStartDocument()
 
#set XSL statements
$PItext = "type='text/xsl' href='style.xsl'"
$xmlWriter.WriteProcessingInstruction("xml-stylesheet", $PItext)
 
$XmlWriter.WriteComment('List of SKC Alma Users')
$xmlWriter.WriteStartElement('users')

foreach ($row in $DataSet.Tables[0].Rows)
{
	$xmlWriter.WriteStartElement('user')
		$xmlWriter.WriteStartElement('record_type')
			$xmlWriter.WriteAttributeString('desc','Public')
			$xmlWriter.WriteRaw('PUBLIC')
		$xmlWriter.WriteEndElement() #Record type
		
		$xmlWriter.WriteElementString('primary_id', $row.id_num)
		$xmlWriter.WriteElementString('first_name', $row.first_name)
		$xmlWriter.WriteElementString('middle_name', $row.middle_name)
		$xmlWriter.WriteElementString('last_name', $row.last_name)
		write-host $row.id_num + " " + $row.last_name
		$fullname = $row.first_name +  $row.last_name
		$xmlWriter.WriteElementString('full_name', $fullname)
		$xmlWriter.WriteStartElement('user_group')
			$xmlWriter.WriteAttributeString('desc',$row.grp_cde)
			$xmlWriter.WriteRaw($row.grp_cde)
		$xmlWriter.WriteEndElement() #user_group
		
		$xmlWriter.WriteStartElement('preferred_language')
			$xmlWriter.WriteAttributeString('desc','English')
			$xmlWriter.WriteRaw('en')
		$xmlWriter.WriteEndElement() #preferred_language
		
        $xmlWriter.WriteElementString('expiry_date', $expdteStr)
        $xmlWriter.WriteElementString('purge_date', $purgeDteStr)

		$xmlWriter.WriteStartElement('account_type')
			$xmlWriter.WriteAttributeString('desc','External')
			$xmlWriter.WriteRaw('EXTERNAL')
		$xmlWriter.WriteEndElement() #account_type
		
		$xmlWriter.WriteElementString('external_id', 'SIS')
		$xmlWriter.WriteStartElement('status')
			$xmlWriter.WriteAttributeString('desc','Active')
			$xmlWriter.WriteRaw('ACTIVE')
		$xmlWriter.WriteEndElement() #status
		
		$xmlWriter.WriteStartElement('user_identifiers')
			$xmlWriter.WriteStartElement('user_identifier')
				$xmlWriter.WriteAttributeString('segment_type','External')
				$xmlWriter.WriteStartElement('id_type')
					$xmlWriter.WriteAttributeString('desc','LDAP User')
					$xmlWriter.WriteRaw('INST_ID')
				$xmlWriter.WriteEndElement() #id_type
				$xmlWriter.WriteElementString('value', $row.web_login)
				$xmlWriter.WriteElementString('status', 'ACTIVE')
			$xmlWriter.WriteEndElement() #user_identifier
		$xmlWriter.WriteEndElement() #user_identifiers

		
		$xmlWriter.WriteStartElement('contact_info')
			#addresses node
			$xmlWriter.WriteStartElement('addresses')
				$xmlWriter.WriteStartElement('address')
					$xmlWriter.WriteAttributeString('preferred','true')
					$xmlWriter.WriteAttributeString('segment_type','External')
					$xmlWriter.WriteElementString('line1', $row.ADDR_LINE_1)
					$xmlWriter.WriteElementString('city', $row.city)
					$xmlWriter.WriteElementString('state_province', $row.state)
					$xmlWriter.WriteElementString('postal_code', $row.zip)
					$xmlWriter.WriteElementString('country', "")
					$xmlWriter.WriteElementString('address_note', "")
					$xmlWriter.WriteStartElement('address_types')
						$xmlWriter.WriteStartElement('address_type')
							$xmlWriter.WriteAttributeString('desc','Home')
							$xmlWriter.WriteRaw('home')
						$xmlWriter.WriteEndElement() #Address_type
					$xmlWriter.WriteEndElement() #Address_types
				$xmlWriter.WriteEndElement() #address
			$xmlWriter.WriteEndElement() #addresses
			
			#email nodes
			$xmlWriter.WriteStartElement('emails')
				$xmlWriter.WriteStartElement('email')
					$xmlWriter.WriteAttributeString('preferred','true')
					$xmlWriter.WriteAttributeString('segment_type','External')
					$xmlWriter.WriteElementString('email_address', $row.email_address)
					$xmlWriter.WriteStartElement('email_types')
						$xmlWriter.WriteStartElement('email_type')
							$xmlWriter.WriteAttributeString('desc','School')
							$xmlWriter.WriteRaw('school')
						$xmlWriter.WriteEndElement() #email_type
					$xmlWriter.WriteEndElement() #email_types
				$xmlWriter.WriteEndElement() #email
			$xmlWriter.WriteEndElement() #emails
			
			#phones nodes
            if ($row.phone -ne 0)
            {
			    $xmlWriter.WriteStartElement('phones')
				    $xmlWriter.WriteStartElement('phone')
					    $xmlWriter.WriteAttributeString('preferred','true')
					    $xmlWriter.WriteAttributeString('preferred_sms','false')
					    $xmlWriter.WriteAttributeString('segment_type','External')
					    $xmlWriter.WriteElementString('phone_number', $row.phone)
					    $xmlWriter.WriteStartElement('phone_types')
						    $xmlWriter.WriteStartElement('phone_type')
							    $xmlWriter.WriteAttributeString('desc','Home')
							    $xmlWriter.WriteRaw('home')
						    $xmlWriter.WriteEndElement() #phone_type
					    $xmlWriter.WriteEndElement() #phone_types
				    $xmlWriter.WriteEndElement() #phone
			    $xmlWriter.WriteEndElement() #phones
            }
		$xmlWriter.WriteEndElement() #contact_info			
	$xmlWriter.WriteEndElement() #user
				
	
}

$xmlWriter.WriteEndElement() #users

$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()

#notepad $XMLPath

# NOTE: You need to install PowerShell Community Extensions for Write-Zip
# Available at https://pscx.codeplex.com/releases/view/133199
Write-Zip $XMLPath "C:\AlmaUpload\skc_alma.zip"

