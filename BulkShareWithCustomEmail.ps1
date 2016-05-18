[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)][string]$url, 
    [Parameter(Mandatory=$true)][string]$domain, 
    [Parameter(Mandatory=$true)][string]$groupTitle,
    [Parameter(Mandatory=$true)][string]$csvFile,
    [Parameter(Mandatory=$true)][string]$csomPath
)

Set-Strictmode -Version 1

$psCredentials = Get-Credential
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($psCredentials.UserName, $psCredentials.Password)

$usersCSV = Import-CSV $csvFile

# the path here may need to change if you used e.g. C:\Lib.. 
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.Runtime.dll" 

# connect/authenticate to SharePoint Online and get ClientContext object.. 
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$clientContext.Credentials = $spoCredentials 

If (!$clientContext.ServerObjectIsNull.Value) { 
    Write-Host "Connected to SharePoint Online site collection: " $url -ForegroundColor Green        
                
    $web = $clientContext.Web
    $clientContext.Load($web)
    $clientContext.Load($web.SiteGroups)
    $clientContext.ExecuteQuery()

    $myGroups = $web.SiteGroups | ? { $_.Title -eq $groupTitle }
    If ($myGroups.Count -eq 0) { 
        Write-Error "Group, $groupTitle, was not found." 
    } Else {
        $myGroups | % {
            $groupNumber = $_.Id
            Write-Host "Found ID for `"$groupTitle`": " $groupNumber -ForegroundColor Green        
            $usersCSV | % {
                $email = $_.Email
                Write-Host "Inviting user: " $email -ForegroundColor Green        
                $peoplePickerValue = "[{`"Key`":`"$email`",`"Description`":`"$email`",`"DisplayText`":`"$email`",`"EntityType`":`"`",`"ProviderDisplayName`":`"`",`"ProviderName`":`"`",`"IsResolved`":true,`"EntityData`":{`"Email`":`"$email`",`"SIPAddress`":`"$email`",`"SPUserID`":`"$email`",`"AccountName`":`"$email`",`"PrincipalType`":`"UNVALIDATED_EMAIL_ADDRESS`"},`"MultipleMatches`":[],`"AutoFillKey`":`"$email`",`"AutoFillDisplayText`":`"$email`",`"AutoFillSubDisplayText`":`"`",`"AutoFillTitleText`":`"$email\n$email`",`"DomainText`":`"$domain`",`"Resolved`":true}]"
                $sharingResult = [Microsoft.SharePoint.Client.Web]::ShareObject($clientContext, $url, $peoplePickerValue, "group:$groupNumber", $groupNumber, $false, $false, $false, "", "")
                $clientContext.Load($sharingResult)
                $clientContext.ExecuteQuery()

                Write-Host "Emailing user: " $email -ForegroundColor Green        
                $invitationLink = $sharingResult.InvitedUsers[0].InvitationLink
                $todaysDate = Get-Date -Format D
                $emailSubject = "Test subject"
                $emailBody = "<h3 style=`"color: red`">Test HTML email</h3><a href=`"$invitationLink`">Click this link to accept the invitation.</a>"

                Send-MailMessage -To $email -From $username -Subject $emailSubject -Body $emailBody -BodyAsHtml -SmtpServer smtp.office365.com -UseSsl -Credential $psCredentials -Port 587
            }
        }
    }
}