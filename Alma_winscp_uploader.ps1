# Load WinSCP .NET assembly
Add-Type -Path (Join-Path $PSScriptRoot "WinSCPnet.dll")

# Set up session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = "weblib.lib.umt.edu"
    UserName = "*******"
    Password = "*******"
    SshHostKeyFingerprint = "*********"
}

$session = New-Object WinSCP.Session

try
{
    # Connect
    $session.Open($sessionOptions)

    # Transfer files
    $session.PutFiles("C:\AlmaUpload\skc_alma.zip", "/home/almaskc/alma/synchronize/*").Check()
}
finally
{
    $session.Dispose()
}
