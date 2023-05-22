# 2017-6-12
# Al Anderson
#
# This PowerShell script is queries the Jenzabar tmseprd db, gets the current instructor and student section enrollments, and creates a file for upload to Schoology
# This runs automatically from a Windows task
# The upload process is a seperate task
#
# Changelog
# 2017-6-28 ala
# Changed SQL for faculty to read from faculty_load table instead of schedule
# 2021-5-5 l
# Heavily modified script for BrightSpace

$PathToScript = split-path -parent $MyInvocation.MyCommand.Path

$termyearPath = $PathToScript + "\term_years.txt"
$termYears = Import-Csv $termyearPath

$term = $termYears.Term
$year = $termYears.Year


function GetTermYearData ($termYear)
{

    $term = $termYear.Term
    $year = $termYear.Year

    $extractFile = $PathToScript + "\temp_csv_files\enrollments-from-db-$year-$term.csv"

	$server = "x.x.x.x"
	$database = "tmseprd"

	
	$query = @"
select distinct
    'enrollment' as 'type',
    'UPDATE' as 'action',
    FACULTY_LOAD_TABLE_V.INSTRCTR_ID_NUM as 'child_code', --INSTRUCTORSOURCEID
    'Instructor' AS 'role', -- Added for Schoology Extract
    '$term-$year-' + left(section_master_v.crs_comp1, 4) + left(section_master_v.crs_comp2, 3)+ '-' + 
        left(replace(SECTION_MASTER_V.crs_comp3,' ', ''),1) + '-' + left(section_master_v.crs_comp4, 2) + '-' + left(section_master_v.crs_comp5, 2) as 'parent_code'


    FROM FACULTY_LOAD_TABLE_V LEFT OUTER JOIN name_master ON FACULTY_LOAD_TABLE_V.INSTRCTR_ID_NUM = name_master.id_num,   
         catalog_master,   
         section_master_v,   
         section_master  
      WHERE ( FACULTY_LOAD_TABLE_V.crs_cde = section_master_v.crs_cde ) and  
         ( FACULTY_LOAD_TABLE_V.yr_cde = section_master_v.yr_cde ) and  
         ( FACULTY_LOAD_TABLE_V.trm_cde = section_master_v.trm_cde ) and  
         ( FACULTY_LOAD_TABLE_V.yr_cde = section_master.yr_cde ) and  
         ( FACULTY_LOAD_TABLE_V.trm_cde = section_master.trm_cde ) and  
         ( FACULTY_LOAD_TABLE_V.crs_cde = section_master.crs_cde ) and  
         ( left(section_master_v.crs_cde, 11) = catalog_master.crs_cde ) and  
         (  ( FACULTY_LOAD_TABLE_V.yr_cde = '$year' ) and  
            ( FACULTY_LOAD_TABLE_V.trm_cde = '$term' ) ) and
			section_master.CRS_CANCEL_FLG <> 'Y' and
            name_master.id_num <> 663
UNION ALL

select  distinct 
    'enrollment' as 'type',
    CASE
		WHEN STUDENT_CRS_HIST.ADD_FLAG = 'A' and student_crs_hist.drop_flag is null
			THEN 'UPDATE'
			ELSE 'DELETE'
			END as 'action',
    --'UPDATE' as 'action',
    name_master.id_num AS 'child_code',
	'Learner' AS 'role',
    '$term-$year-' + left(section_master_v.crs_comp1, 4) + left(section_master_v.crs_comp2, 3)+ '-' + 
        left(replace(SECTION_MASTER_V.crs_comp3,' ', ''),1) + '-' + left(section_master_v.crs_comp4, 2) + '-' + left(section_master_v.crs_comp5, 2) as 'parent_code'

    FROM name_master,   
         section_schedules_v,
			section_master_v, 
			section_master,  
         student_crs_hist  
   WHERE 
 ( section_schedules_v.crs_cde = section_master_v.crs_cde ) and  
         ( section_schedules_v.yr_cde = section_master_v.yr_cde ) and  
         ( section_schedules_v.trm_cde = section_master_v.trm_cde ) and  
         ( section_schedules_v.yr_cde = section_master.yr_cde ) and  
         ( section_schedules_v.trm_cde = section_master.trm_cde ) and  
         ( section_schedules_v.crs_cde = section_master.crs_cde ) and
( student_crs_hist.id_num = name_master.id_num ) and  
         ( student_crs_hist.yr_cde = section_schedules_v.yr_cde ) and  
         ( student_crs_hist.trm_cde = section_schedules_v.trm_cde ) and  
--section_master_v.institut_div_cde IN ('CT', 'GE', 'PC','AH') and NOT (rtrim(section_schedules_v.crs_comp1) IN ('ASL','DAE','DEH','GLG','GHY','NAP','NUR','PNC','SPA','EGR','PHY') and rtrim(section_schedules_v.crs_comp3) = 'L') and 
--	rtrim(section_schedules_v.crs_comp1) NOT IN ('ORT') and
         ( student_crs_hist.crs_cde = section_schedules_v.crs_cde ) and  
         ( ( student_crs_hist.yr_cde = '$year' ) AND  
         ( student_crs_hist.trm_cde = '$term' ) ) AND 
	  (student_crs_hist.transaction_sts = 'C' or student_crs_hist.transaction_sts = 'D' or student_crs_hist.transaction_sts = 'H') AND
	  name_master.email_address is not null
--	  section_master_v.crs_comp4 NOT LIKE 'P%%' AND
--	  section_master_v.crs_comp4 NOT LIKE 'C%%'

"@


	$connectionTemplate = "Data Source={0};User ID=******;Password=******;Initial Catalog={1};"
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
	$DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Select-Object -Skip 1 | out-file $extractFile -Force -Encoding ascii
}



# Call the function to actually get the data and write to a file named for year and term

    foreach ($ty in $termYears) {
		Write-Host "$ty"
        GetTermYearData($ty)
    }



# dump the data to a csv
# http://technet.microsoft.com/en-us/library/ee176825.aspx
# $DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Export-Csv $extractFile
#$DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Select-Object -Skip 1 | out-file $extractFile -Force -Encoding ascii


# Combine enrollment-from-db.csv with enrollment-manual.csv into final enrollment.csv
# Add manual enrollments to enrollment-manual.csv file in the specified format
#Get-Content $extractFile, $manualEnrollmentFile | Set-Content $finalEnrollmentFile

# Need to use a header file to write the column headings
$headerEnrollmentFile = $PathToScript + "\headers\enrollments-header.csv"

# $manualEnrollmentFile = $PathToScript + "\enrollments-manual.csv"

# tempFile is used to update final enrollment file
$tempFile = $PathToScript + "\temp_csv_files\temp.csv"

$finalEnrollmentFile = $PathToScript + "\current_csv_files\enrollments_final.csv"


Get-Content $headerEnrollmentFile | Set-Content $finalEnrollmentFile

foreach ($ty in $termYears) {
	$extractFile = $PathToScript + "\temp_csv_files\enrollments-from-db-$($ty.Year)-$($ty.Term).csv"
	Get-Content $finalEnrollmentFile | Set-Content $tempFile
	Get-Content $tempFile, $extractFile | Set-Content $finalEnrollmentFile
}

#Copy final file over to Dated folder

$brightspace_csv_folder = $PathToScript + "\brightspace_csv_files\"
$new_folder = Get-Date -Format "yyyy_MM_dd_HHmm"

$brightspace_csv_new_folder_name = $brightspace_csv_folder + $new_folder

if(-Not (Test-Path $brightspace_csv_new_folder_name)){
    New-Item -Path $brightspace_csv_folder -Name $new_folder -ItemType "directory"
}

$new_folder_name = $brightspace_csv_new_folder_name + "\enrollments_$new_folder.csv"

Copy-Item $finalEnrollmentFile -Destination $new_folder_name

