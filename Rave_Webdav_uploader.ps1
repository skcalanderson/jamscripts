#Setting WebDAV Connection Details

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$UserName = "*********";
$Password = "*********"; 

$WebCredential = New-Object System.Net.NetworkCredential($UserName, $Password);
$FilePath = "C:\RAVE-Uploader\rave-people.csv";
$raveURL = "https://************";

 

#Uploading File
Write-Output "Uploading file $FilePath to $raveURL";
$WebClient = New-Object System.Net.WebClient;
$WebClient.Credentials = $WebCredential;
$WebClient.UploadFile($raveURL, 'PUT', $FilePath);

 

Write-Output "Upload complete";