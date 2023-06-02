# 2016-10-19
# Al Anderson
#
# This PowerShell script runs a query on the Jenzabar tmseprd dbr, retrieves current studens and employees, and creates a file for upload to RAVE
#
# If the above fails due to
# "cannot be loaded because the execution of scripts is disabled on this system"
# run this command within Powershell
# Set-ExecutionPolicy RemoteSigned
# See also http://technet.microsoft.com/en-us/library/ee176949.aspx


$server = "x.x.x.x"
$database = "tmseprd"
$query = "SELECT id_num,x_first,x_last,email_address,x_gender,x_mobile_phone,x_7to16,x_landline_phone1,x_18,x_landline_phone2 FROM skc_rave_extract_emplye_stdnt_v"

# powershell raw/verbatim strings are ugly
# Update this with the actual path where you want data dumped
$extractFile = @"
C:\RAVE-Uploader\rave-people.csv
"@

# If you have to use users and passwords, my condolences
# Don't be dumb like me and use at least Base64 to hide the password
$connectionTemplate = "Data Source={0};User ID=******;Password=*******;Initial Catalog={1};"
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

# dump the data to a csv
# http://technet.microsoft.com/en-us/library/ee176825.aspx
# $DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Export-Csv $extractFile
$DataSet.Tables[0] | ConvertTo-Csv -NoTypeInformation | %{$_ -replace '"', ''} | Select-Object -Skip 1 | out-file $extractFile -Force -Encoding ascii
