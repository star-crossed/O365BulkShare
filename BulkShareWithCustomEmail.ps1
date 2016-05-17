# replace these details (also consider using Get-Credential to enter password securely as script runs).. 
$username = "REDACTED" 
$password = "REDACTED" 
$url = "REDACTED"
$domain = "REDACTED.sharepoint.com"
$groupNumber = REDACTED

$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force 
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword) 
$psCredentials = New-Object System.Management.Automation.PSCredential($username, $securePassword)

$usersCSV = Import-CSV "c:\Users\pchoquette\Desktop\users.csv"

# the path here may need to change if you used e.g. C:\Lib.. 
Add-Type -Path "C:\Users\pchoquette\Source\Repos\PnP-Sites-Core\Assemblies\16.1\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "C:\Users\pchoquette\Source\Repos\PnP-Sites-Core\Assemblies\16.1\Microsoft.SharePoint.Client.Runtime.dll" 

# connect/authenticate to SharePoint Online and get ClientContext object.. 
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$clientContext.Credentials = $spoCredentials 

if (!$clientContext.ServerObjectIsNull.Value) { 
        Write-Host "Connected to SharePoint Online site collection: " $url -ForegroundColor Green 
        
        $usersCSV | % {
            $email = $_.Email
            $peoplePickerValue = "[{`"Key`":`"$email`",`"Description`":`"$email`",`"DisplayText`":`"$email`",`"EntityType`":`"`",`"ProviderDisplayName`":`"`",`"ProviderName`":`"`",`"IsResolved`":true,`"EntityData`":{`"Email`":`"$email`",`"SIPAddress`":`"$email`",`"SPUserID`":`"$email`",`"AccountName`":`"$email`",`"PrincipalType`":`"UNVALIDATED_EMAIL_ADDRESS`"},`"MultipleMatches`":[],`"AutoFillKey`":`"$email`",`"AutoFillDisplayText`":`"$email`",`"AutoFillSubDisplayText`":`"`",`"AutoFillTitleText`":`"$email\n$email`",`"DomainText`":`"$domain`",`"Resolved`":true}]"
            $sharingResult = [Microsoft.SharePoint.Client.Web]::ShareObject($clientContext, $url, $peoplePickerValue, "group:$groupNumber", $groupNumber, $false, $false, $false, "", "")
            $clientContext.Load($sharingResult)
            $clientContext.ExecuteQuery()

            $invitationLink = $sharingResult.InvitedUsers[0].InvitationLink
            $emailSubject = "Test subject"
            $emailBody = "<h3 style=`"color: red`">Test HTML email</h3><a href=`"$invitationLink`">Click this link to accept the invitation.</a>"
            Send-MailMessage -To $email -From $username -Subject $emailSubject -Body $emailBody -BodyAsHtml -SmtpServer smtp.office365.com -UseSsl -Credential $psCredentials -Port 587
        }
}