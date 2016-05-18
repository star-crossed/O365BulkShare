[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, HelpMessage="This is the URL to the SharePoint Online site where you are inviting users.")][string]$Url, 
    [Parameter(Mandatory=$true, HelpMessage="This is the display name of the group on your SharePoint Online site where users will be added.")][string]$GroupTitle,
    [Parameter(Mandatory=$true, HelpMessage="This is the path to the CSV file that has a single column, Email, which contains each email address to be invited.")][string]$CSVFile,
    [Parameter(Mandatory=$true, HelpMessage="This is the path to the DLLs for CSOM.")][string]$CSOMPath
)

Set-Strictmode -Version 1

$psCredentials = Get-Credential
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($psCredentials.UserName, $psCredentials.Password)

$usersCSV = Import-CSV $CSVFile

$domain = ([System.Uri]$Url).Host

# the path here may need to change if you used e.g. C:\Lib.. 
Add-Type -Path "$CSOMPath\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$CSOMPath\Microsoft.SharePoint.Client.Runtime.dll" 

# connect/authenticate to SharePoint Online and get ClientContext object.. 
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($Url) 
$clientContext.Credentials = $spoCredentials 

If ($clientContext.ServerObjectIsNull.Value) { 
    Write-Error "Could not connect to SharePoint Online site collection: $Url"
} Else {
    Write-Host "Connected to SharePoint Online site collection: " $Url -ForegroundColor Green        
                
    $web = $clientContext.Web
    $clientContext.Load($web)
    $clientContext.Load($web.SiteGroups)
    $clientContext.ExecuteQuery()

    $myGroups = $web.SiteGroups | ? { $_.Title -eq $GroupTitle }
    If ($myGroups.Count -eq 0) { 
        Write-Error "Group, $GroupTitle, was not found." 
    } Else {
        $myGroups | % {
            $groupNumber = $_.Id
            Write-Host "Found ID for `"$GroupTitle`": " $groupNumber -ForegroundColor Green        
            $usersCSV | % {
                $email = $_.Email
                Write-Host "Inviting user: " $email -ForegroundColor Green        
                $peoplePickerValue = "[{`"Key`":`"$email`",`"Description`":`"$email`",`"DisplayText`":`"$email`",`"EntityType`":`"`",`"ProviderDisplayName`":`"`",`"ProviderName`":`"`",`"IsResolved`":true,`"EntityData`":{`"Email`":`"$email`",`"SIPAddress`":`"$email`",`"SPUserID`":`"$email`",`"AccountName`":`"$email`",`"PrincipalType`":`"UNVALIDATED_EMAIL_ADDRESS`"},`"MultipleMatches`":[],`"AutoFillKey`":`"$email`",`"AutoFillDisplayText`":`"$email`",`"AutoFillSubDisplayText`":`"`",`"AutoFillTitleText`":`"$email\n$email`",`"DomainText`":`"$myDomain`",`"Resolved`":true}]"
                $sharingResult = [Microsoft.SharePoint.Client.Web]::ShareObject($clientContext, $Url, $peoplePickerValue, "group:$groupNumber", $groupNumber, $false, $false, $false, "", "")
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