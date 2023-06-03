# Script for Using PowerShell 7 to Generate CSV & XML from Jenzabar
JAM 2023 PowerShell Scripts

These are the basic scripts talked about at JAM 2023

## Security of SQL Login
If you can, it is best to use a trusted connection for the SQL connection, barring that, the next
best thing it seems is to protect the credentials in some fashion. This article has some great information: [SECURELY STORING CREDENTIALS WITH POWERSHELL](https://metisit.com/blog/securely-storing-credentials-with-powershell/). This Bitbucket [repo](https://bitbucket.org/metisit/credentialmanager/src/master/) is for a library that uses the Windows Data Protection API (DPAPI). This work is by **Theo Hardendood** who authored the above blog.
