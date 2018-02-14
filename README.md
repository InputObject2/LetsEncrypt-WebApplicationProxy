# LetsEncrypt-WebApplicationProxy
A script based on the Lone-Coder's LetsEncryptWinSimple to automatically do certificate renewals on a Microsoft Web Application Proxy.

## Synopsis
This script allows automatically setting-up and renewing all the WebApplications with LetsEncrypt certificates.

## Description
This script will install a new event log source and log events in "Application" log. It will also install a scheduled
task to do the renewals each day at 3h AM. It gets the list of web applications, looks in the "Personnal" certificate
repository on the computer account to make sure that the certs are still valid.
If a cert has to be renewed or created, it gets added with letsEncrypt.

## Instructions
To get started, you have to :  
- Download LetsEncryptWinSimple (https://github.com/Lone-Coder/letsencrypt-win-simple/releases).  
- Extract the zip to "C:\Program Files\LetsEncryptWinSimple\" or update the variable LetsEncryptPath to point to it.  
- Copy this script in "C:\LetsEncrypt\" or update the variable to point to it.  
- Install IIS on the server, no need to configure it or anything.  
- Run this script once manually.  

To add WAP applications, you need to set them manually in the WAP console. Just assign them a BS cert, this script will take care
of making sure it is letsEncrypt ready next time it runs.   

Be aware, there's a 20 certs/week rate limit so there might be a bit of pain if you have a lot of WAP applications.
