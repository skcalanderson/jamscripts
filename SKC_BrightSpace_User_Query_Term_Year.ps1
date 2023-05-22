# 2017-1-12
# Al Anderson
#
# This PowerShell script is queries the Jenzabar tmseprd db, gets the current instructors and students, and creates a file for upload to Schoology
# This runs automatically from a Windows task
# The upload process is a seperate task
#
# Changelog
# 2017-6-28 ala
# Changed SQL for faculty to read from faculty_load table instead of schedule
# 2021-5-6 ala
# Modified for BrightSpace CSV export

$PathToScript = split-path -parent $MyInvocation.MyCommand.Path
$termyearPath = $PathToScript + "\term_years.txt"
$termYears = Import-Csv $termyearPath

$temp = ""

foreach ($h in $termYears) {
    $temp = $temp, "'$($h.Year)$($h.Term)'" -join ","
}

$queryYearTermString = $temp.Substring(1)

Write-Host  $queryYearTermString


$server = "x.x.x.x"
$database = "tmseprd"

# Major change to SQL script
# 2017-4-4
# Needed to not build the email address instead just copy it
# Also removed all employees from having student user records created.
# If they are in a course as a student then they will be in the enrollment file as a student. 
$query = @"
		SELECT DISTINCT
        'user' as 'type',
        'UPDATE' as 'action',

		LOWER(
			CASE
				WHEN(CHARINDEX('@skc.edu',nm.email_address)) > 0
				THEN SUBSTRING(nm.email_address,0,charindex('@', nm.EMAIL_ADDRESS))
				WHEN(CHARINDEX('@student.skc.edu',nm.email_address)) > 0
				THEN REPLACE( SUBSTRING(nm.email_address,0,charindex('@', nm.EMAIL_ADDRESS)), ' ', '') 
				ELSE Replace(nm.first_name + nm.last_name, ' ', '')
			END)
		AS 'username',
        sm.id_num as 'org_defined_id',
        nm.first_name AS 'first_name',
        nm.last_name AS 'last_name',
		'' as 'password',
		'1' as 'is_active',
		'Learner' AS 'role_name',
		LOWER(
			CASE
				WHEN (CHARINDEX('@student.skc.edu',nm.email_address)) >0
					THEN nm.EMAIL_ADDRESS
				WHEN (CHARINDEX('@skc.edu',nm.email_address)) > 0
					THEN nm.EMAIL_ADDRESS
					ELSE nm.FIRST_NAME + nm.LAST_NAME + '@student.skc.edu'
			END)
		AS 'email',
		'' as 'relationships',
		'' as 'pref_first_name',
		'' as 'pref_last_name' 

	FROM name_master AS nm
	JOIN student_master sm ON sm.id_num=nm.id_num
	JOIN student_crs_hist sch ON sch.id_num=sm.id_num
	LEFT JOIN STUDENT_VIEW sv ON sv.ID_NUM=sch.ID_NUM 
	WHERE 
	CAST(sch.YR_CDE AS VARCHAR(4))+ CAST (sch.TRM_CDE AS VARCHAR(4)) IN ($queryYearTermString) AND
	sch.transaction_sts = 'C' 
	AND nm.EMAIL_ADDRESS is not null

UNION

SELECT DISTINCT

'user' as 'type',
'UPDATE' as 'action',
LOWER(
	CASE
		WHEN(CHARINDEX('@skc.edu',nm.email_address)) = 0
			THEN LEFT(nm.first_name,1) + REPLACE(REPLACE(REPLACE(LEFT(nm.last_name, 16),'-',''),' ',''),'''','')
		ELSE REPLACE(REPLACE(REPLACE(LEFT(nm.email_address, CHARINDEX('@',nm.email_address)-1),'-',''),' ',''),'''','')
	END)
AS 'username',
ssv.INSTRCTR_ID_NUM AS 'org_defined_id', --InstructorID
nm.first_name AS 'first_name',
nm.last_name AS 'last_name',
'' as 'password',
'1' as 'is_active',
'Instructor' AS 'role_name',
LOWER( nm.email_address) AS 'email',
'' as 'relationships',
'' as 'pref_first_name',
'' as 'pref_last_name'

FROM FACULTY_LOAD_TABLE_V AS ssv, name_master AS nm, section_master_v AS smv

WHERE 
(ssv.INSTRCTR_ID_NUM = nm.id_num) AND
(ssv.crs_cde = smv.crs_cde) AND
CAST(SSV.YR_CDE AS VARCHAR(4))+ CAST (SSV.TRM_CDE AS VARCHAR(4)) IN ($queryYearTermString) AND
smv.section_sts != 'C' AND
-- Exclude Staff
ssv.INSTRCTR_ID_NUM NOT IN (999999,663)
AND nm.EMAIL_ADDRESS is not null
"@


# powershell raw/verbatim strings are ugly
# Update this with the actual path where you want data dumped
$extractFile = $PathToScript + "\temp_csv_files\users.csv"

# If you have to use users and passwords, my condolences
$connectionTemplate = "Data Source={0};User ID=*******;Password=*******;Initial Catalog={1};"
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

#foreach ($row in $DataSet.Tables[0].Rows)
#{
#    Write-Host $row.id_num
#}

# dump the data to a csv
# http://technet.microsoft.com/en-us/library/ee176825.aspx
# $DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Export-Csv $extractFile
#$DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Select-Object -Skip 1 | out-file $extractFile -Force -Encoding ascii
$DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''}  | out-file $extractFile -Force -Encoding ascii

# Need to use a header file to write the column headings
#$headerUserFile = $PathToScript + "\user-header-brightspace.csv"

# $previousUserFileName = $PathToScript + "\users_previous.csv"

#$previousUserFile = Import-Csv $previousUserFileName


$finalUsersFile = $PathToScript + "\current_csv_files\users_final.csv"

Get-Content $extractFile | Set-Content $finalUsersFile


#Copy final file over to Dated folder

$brightspace_csv_folder = $PathToScript + "\brightspace_csv_files\"
$new_folder = Get-Date -Format "yyyy_MM_dd_HHmm"

$brightspace_csv_new_folder_name = $brightspace_csv_folder + $new_folder

if(-Not (Test-Path $brightspace_csv_new_folder_name)){
    New-Item -Path $brightspace_csv_folder -Name $new_folder -ItemType "directory"
}

$new_folder_name = $brightspace_csv_new_folder_name + "\users_$new_folder.csv"

Copy-Item $finalUsersFile -Destination $new_folder_name


#$compareParams = @{
#	ReferenceObject = $finalUsersFile
#	DifferenceObject = $previousUserFile
#}


#$comparison = Compare-Object @compareParams -Property 'org_defined_id'
#$userDifferenceFileName = $PathToScript + '\users_differences.csv' 

#$comparison | Export-Csv  $userDifferenceFileName -NoTypeInformation