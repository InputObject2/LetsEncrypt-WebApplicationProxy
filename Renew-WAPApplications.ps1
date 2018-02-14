
<#
.SYNOPSIS

This script allows automatically setting-up and renewing all the WebApplications with LetsEncrypt certificates.

.DESCRIPTION

This script will install a new event log source and log events in "Application" log. It will also install a scheduled
task to do the renewals each day at 3h AM. It gets the list of web applications, looks in the "Personnal" certificate
repository on the computer account to make sure that the certs are still valid.

If a cert has to be renewed or created, it gets added with letsEncrypt.

.NOTES

Author : PulseBright <30133702+PulseBright@users.noreply.github.com>

.HELP

To get started, you have to :
1- Download LetsEncryptWinSimple (https://github.com/Lone-Coder/letsencrypt-win-simple/releases).
2- Extract the zip to "C:\Program Files\LetsEncryptWinSimple\" or update the variable LetsEncryptPath to point to it.
3- Copy this script in "C:\LetsEncrypt\" or update the variable to point to it.
4- Install IIS on the server, no need to configure it or anything.
5- Run this script once manually.

To add WAP applications, you need to set them manually in the WAP console. Just assign them a BS cert, this script will take care
of making sure it is letsEncrypt ready next time it runs. 

Be aware, there's a 20 certs/week rate limit so there might be a bit of pain if you have a lot of WAP applications.

#>

# Parameters intended to send a daily email with the certificate status for the server.
$subject = "Certificate Renewal - $((get-date).tostring("dd/MM/yyyy"))" 
$smtphost = "exchange.domain.com" 
$from = "automatic.renewal@domain.com" 
$to = "support@domain.com"
$port = 587

$letsEncryptPath = "C:\Program Files\LetsEncryptWinSimple\letsencrypt-win-simple.v1.9.8.2"
$scriptpath = "C:\LetsEncrypt\Add-WAPApplications.ps1"

try{

<# 
    Doesn't work on 2012 or R2, since I'm using the http -> https feature.
    Maybe it works on 2012 or R2 without setting up "-EnableHTTPRedirect:$false"
    in the Get-WebApplicationProxyApplication calls since 2012R2 doesn't do the http
    redirect by default. 
#>

if([Environment]::OSVersion.Version.Major -lt 10){throw "This script only works on server 2016 and above... sorry !"}

import-module WebAdministration
$ExpiredresultSet = ""
$NotExpiredresultSet = ""

$ExpiredresultSetPlain = ""
$NotExpiredresultSetPlain = ""
$logfile =""

$expiredCount = 0
$NotexpiredCount = 0

# We try to add the eventlog, even if it is already there, it doesn't add it twice.
try{New-EventLog -LogName Application -Source "Scheduled Certificate Renewal" -ErrorAction SilentlyContinue} catch {}

Write-EventLog -LogName "Application"`
                -Source "Scheduled Certificate Renewal"`
                -EntryType Information `
                -EventId 20020 `
                -Message "System is beginning scheduled update for the certificates in the Web Application Proxy's database."

                write-host "System is beginning scheduled update for the certificates in the Web Application Proxy's database."

$applist = Get-WebApplicationProxyApplication

Write-EventLog -LogName "Application"`
                -Source "Scheduled Certificate Renewal"`
                -EntryType Information `
                -EventId 20021 `
                -Message "System has retrieved a list of the applications and the verification process will now be underway."

                write-host "System has retrieved a list of the applications and the verification process will now be underway."

foreach($app in $applist) {

    $fqdn = $app.externalURL.Split("/") | ? {($_ -ne "") -and ($_ -notlike "htt*")}

    #this long line gets the cert from the app name and returns the expiration date. 
    $certExpiry = (get-childitem Cert:\LocalMachine\My\ | `
        select subject,notafter,@{ Label = 'IssuedBy' ;Expression = { $_.GetNameInfo( 'SimpleName', $true ) } } | `
        ? {($_.IssuedBy -like "Let*") -and ($_.Subject -like "CN=$fqdn")} | `
        select -First 1).NotAfter

    #if there's no certificate matching (it's using a bullshit certificate for
    # now, not LetsEncrypt) so we are going to renew it.
    if(!$certExpiry) {$certExpiry = (get-date)}

    #we give ourselves a whole week to troublehoot issues.
    $safetyPeriod = (get-date).AddDays(7)
    
    #We checkup to only renew the certs that are about to expire.
    if(!($safetyPeriod -lt $certExpiry)) {

    

    write-host -ForegroundColor Red "Certificate for site $($app.name) will expire on $($certExpiry.toString("dd/MM/yyyy")). We are renewing it now."

    Write-EventLog -LogName "Application"`
                -Source "Scheduled Certificate Renewal"`
                -EntryType Warning `
                -EventId 20022 `
                -Message "Certificate for application $($app.name) is expired. Time for a replacement"

        # Not really implemented yet, but I wanted to be able to put red cells in the table when there's an error.
        # Need to parse the letsEncrypt log to see what's up.
        $failure = $false
        
        # Stop redirecting HTTP to HTTPS so that we can get the http token.
        Get-WebApplicationProxyApplication -Name $app.name | Set-WebApplicationProxyApplication -EnableHTTPRedirect:$false
             
       

        <#========== A bit of borrowed code. ==========#>
        <#
            The next section (creating the website is taking from Alex Chaika's Create-IISWebsite.ps1 script. 
            You may also notice that it has been very heavily edited/simplified to fit in 10 lines. 

            Author: Alex Chaika Date:   May 16, 2016
            Link : https://4sysops.com/archives/create-iis-websites-with-powershell/
        #>

        $websitename = "$($app.name.substring(0,1).toupper()+$app.name.substring(1).tolower())"
        $RootFSFolder = "C:\inetpub\letsencrypt\$($app.name)\"
        $Hostheaders =  "$fqdn"
        
        $addedWebsite = $false

        if(!(Test-path $RootFSFolder)) {New-Item -ItemType directory -Path $RootFSFolder -Force}
        if(!(Test-Path ("IIS:\AppPools\" + $WebSiteName))) {New-WebAppPool -Name $WebSiteName}
        if(!(Test-Path "IIS:\Sites\$WebSiteName")){ 
            $addedWebsite = $true
            New-Website -Name $WebSiteName -PhysicalPath $RootFSFolder  -ApplicationPool $WebSiteName | Out-Null
        }
        
        if($addedWebsite) {Get-Website | ? {$_.Name -eq $WebSiteName} | % {
            New-WebBinding -Name $_.Name -HostHeader $Hostheaders -IP "*" -Port 80 -Protocol http
            New-WebBinding -Name $_.Name -HostHeader $Hostheaders -IP "*" -Port 9443 -Protocol https} | out-null
            
            #Remove empty bindings (there's one automatically generated with no name).
            Get-WebBinding | ?{($_.bindingInformation).Length -eq 5} | Remove-WebBinding
        }

        # Gotta make sure that bad boy is running.
        Get-Website | ? {$_.Name -eq $WebSiteName} | Start-Website | Out-Null

        <#========== End of borrowed code. ==========#>

        #we remove the cert from the cert store to make sure we don't fudge it up.
        $((get-childitem Cert:\LocalMachine\My\ | `
        select subject,notafter,@{ Label = 'IssuedBy' ;Expression = { $_.GetNameInfo( 'SimpleName', $true ) } } | `
        ? {($_.IssuedBy -like "Let*") -and ($_.Subject -like "CN=$fqdn")} | `
        select -First 1)) | Remove-Item

        #We use letsencrypt to generate a new cert.
        &"$letsEncryptPath\letsencrypt.exe" --plugin "manual" --manualhost $fqdn --manualtargetisiis --webroot "C:\inetpub\letsencrypt\$($app.name)\" --installation "none" --notaskscheduler --certificatestore "My"

        #once done, we remove the website.
        try {Remove-Website -name "$($app.name)"} catch {<#Do nothing, if this fails it means there was no website.#>}

        #Updating the Web app with the certificate.
        try {Set-WebApplicationProxyApplication `
             -ID $app.ID `
             -BackendServerUrl "$($app.backendserverurl)"`
             -ExternalCertificateThumbprint $((get-childitem Cert:\LocalMachine\My\ | select subject,notafter,@{ Label = 'IssuedBy' ;Expression = { $_.GetNameInfo( 'SimpleName', $true ) } } | ? {($_.IssuedBy -like "Let*") -and ($_.Subject -like "CN=$fqdn")} | select -First 1).thumbprint)`
             -ExternalUrl "$($app.externalURL)"`
             -Name "$($app.name.substring(0,1).toupper()+$app.name.substring(1).tolower())"`
             -EnableHTTPRedirect:$true} catch {
                $failure = $true
             }
        
        $ExpiredresultSet += $(if(!$failure){"<tr style=`"background-color:red`">"}else{"<tr>"}) + "<td><strong>$($app.name)</strong></td><td align=`"center`">$(if(!$failure){"Failed"}else{"Renewed"})</td></tr>"
        $ExpiredresultSetPlain += "$($app.name)`t$(if(!$failure){"Failed"}else{"Renewed"})`n"
        $expiredCount++

    } else {
        
         write-host -ForegroundColor Green "Certificate for site $($app.name) is OK."

        $notexpiredresultset += "<tr><td><strong>$($app.name)</strong></td><td align=`"center`" " + $(if(($certExpiry-(get-date)).Days -gt 30){ "style=`"background-color:green`""} else {"style=`"background-color:yellow`""}) + ">$(($certExpiry-(get-date)).Days)</td><td align=`"center`">$($certExpiry.toString("dd/MM/yyyy"))</td></tr>"
        $notExpiredresultSetPlain += "$($app.name)`t$(($certExpiry-(get-date)).Days)`t$($certExpiry.toString("dd/MM/yyyy"))`n"
        $NotexpiredCount++
    }
}
}
catch{

Write-EventLog -LogName "Application"`
                -Source "Scheduled Certificate Renewal"`
                -EntryType Error `
                -EventId 20022 `
                -Message "$($_.Exception.Message)"
}

# This bit is separated into a bunch of lines for ease of reading / editing, I could make it a one liner but it would be disgustingly long.
$message = "<p>The LetsEncrypt certificates for the Web Application Proxy <strong>`"$(hostname)`"</strong> have been inspected.</p>"
$message += "<p>Of the $($expiredCount+$NotexpiredCount) applications inspected, $(if($expiredCount -eq 0){"none"} else {$expiredCount}) were expired"
$message += " and $(if($notexpiredCount -eq 0){"none"} else {$notexpiredCount}) were not.</p>"
$message += "$(if($expiredCount -ne 0){"<p>The following list was expired and then renewed ($expiredCount) :"})"
$message += $(if($expiredCount -ne 0){"<table border=`"1`", style=`"width:80%, table-layout:auto`"><tr><th>Site</th><th>Status</th>"})
$message += $(if($expiredCount -ne 0){"$ExpiredresultSet </table>"})
$message += "$(if($notexpiredCount -ne 0){"<p>The following list was not expired ($notexpiredCount) :</p>"})"
$message += $(if($notexpiredCount -ne 0){"<table border=`"1`", style=`"width:80%, table-layout:auto`"><tr><th>Site</th><th>Days left</th><th>Expires on</th>"})
$message += $(if($notexpiredCount -ne 0){"$NotExpiredresultSet </table>"})

$messagePlain = "The LetsEncrypt certificates for the Web Application Proxy `"$(hostname)`" have been inspected.`n"
$messagePlain += "Of the $($expiredCount+$NotexpiredCount) applications inspected, $(if($expiredCount -eq 0){"none"} else {$expiredCount}) were expired"
$messagePlain += " and $(if($notexpiredCount -eq 0){"none"} else {$notexpiredCount}) were not."
$messagePlain += "$(if($expiredCount -ne 0){"`nThe following list was expired and then renewed ($expiredCount) :`n"})"
$messagePlain += $(if($expiredCount -ne 0){"`nSite`tStatus`n"})
$messagePlain += $(if($expiredCount -ne 0){"$ExpiredresultSetPlain"})
$messagePlain += "$(if($notexpiredCount -ne 0){"`nThe following list was not expired ($notexpiredCount) :`n"})"
$messagePlain += $(if($notexpiredCount -ne 0){"`nSite`tDays left`tExpires`n"})
$messagePlain += $(if($notexpiredCount -ne 0){"$NotExpiredresultSetPlain"})


Write-EventLog -LogName "Application"`
                -Source "Scheduled Certificate Renewal"`
                -EntryType Information `
                -EventId 20022 `
                -Message $messagePlain

Write-host "Event log updated with verification results."



Send-MailMessage -to $to -From "Certificate Renewal Reports <$from>" -Subject $subject -Body $message -BodyAsHtml -SmtpServer $smtphost -Port $port

Write-Host "Email report sent to `"$to`"."

# A quick bit of code to add the scheduled task if it wasn't installed yet.
if(!(Get-ScheduledTask -taskname "Scheduled update all sites")){
    
    #Thanks to the scripting guys. 
    #Link : https://blogs.technet.microsoft.com/heyscriptingguy/2015/01/13/use-powershell-to-create-scheduled-tasks/
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' `
    -Argument "-c $scriptpath"
    $trigger =  New-ScheduledTaskTrigger -Daily -At 3am
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Scheduled update all sites" -Description "Scheduled update all sites" -User SYSTEM -RunLevel Highest
}

write-host "Scheduled task verified."
