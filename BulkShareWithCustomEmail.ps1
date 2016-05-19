[CmdletBinding(DefaultParameterSetName="UseCSV")]
Param(
    [Parameter(Mandatory=$true, Position=0, ParameterSetName="UseCSV", HelpMessage="This is the path to the CSV file that has a single column, Email, which contains each email address to be invited.")]
    [string]$CSVFile,

    [Parameter(Mandatory=$true, Position=0, ParameterSetName="UseEmail", HelpMessage="This is the email address to be invited.")]
    [string]$UserEmail,

    [Parameter(Mandatory=$true, ParameterSetName="UseEmail", HelpMessage="This is the email address to be invited.")]
    [Parameter(Mandatory=$true, ParameterSetName="UseCSV", HelpMessage="This is the path to the CSV file that has a single column, Email, which contains each email address to be invited.")]
    [string]$Url, 

    [Parameter(Mandatory=$true, ParameterSetName="UseEmail", HelpMessage="This is the email address to be invited.")]
    [Parameter(Mandatory=$true, ParameterSetName="UseCSV", HelpMessage="This is the path to the CSV file that has a single column, Email, which contains each email address to be invited.")]
    [string]$GroupTitle,

    [Parameter(Mandatory=$false, ParameterSetName="UseEmail", HelpMessage="This is the email address to be invited.")]
    [Parameter(Mandatory=$false, ParameterSetName="UseCSV", HelpMessage="This is the path to the CSV file that has a single column, Email, which contains each email address to be invited.")]
    [string]$CSOMPath,

    [Parameter(Mandatory=$true, ParameterSetName="UseCSV", HelpMessage="This is the number of email addresses to include in one batch.")]
    [int]$BatchAmount,

    [Parameter(Mandatory=$true, ParameterSetName="UseCSV", HelpMessage="This is the amount of seconds to wait between batches.")]
    [int]$BatchInterval
)

Set-Strictmode -Version 1

If ($CSOMPath -eq $null -or $CSOMPath -eq "") { $CSOMPath = "." }

Add-Type -Path "$CSOMPath\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$CSOMPath\Microsoft.SharePoint.Client.Runtime.dll" 

If ($PSCmdlet.ParameterSetName -eq "UseCSV") {
    $usersCSV = Import-CSV $CSVFile
} Else {
    $BatchAmount = 1
    $BatchInterval = 0
    $usersCSV = @(@{
        "Email"="$UserEmail";
    })
}

$domain = ([System.Uri]$Url).Host

# connect/authenticate to SharePoint Online and get ClientContext object.. 
$psCredentials = Get-Credential
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($psCredentials.UserName, $psCredentials.Password)
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
            $i = 0
            $usersCSV | % { 
                Add-Member -InputObject $_ -MemberType NoteProperty -Name "Row" -Value $i; $i++ 
            } 
            $usersCSV | Group-Object { 
                [System.Math]::Truncate($_.Row / $BatchAmount) 
            } | % {
                Write-Host "--- Start of Batch ---"
                $_.Group | % { 
                    $email = $_.Email

                    Write-Host "Inviting user: " $email -ForegroundColor Green     
                    $peoplePickerValue = "[{`"Key`":`"$email`",`"Description`":`"$email`",`"DisplayText`":`"$email`",`"EntityType`":`"`",`"ProviderDisplayName`":`"`",`"ProviderName`":`"`",`"IsResolved`":true,`"EntityData`":{`"Email`":`"$email`",`"SIPAddress`":`"$email`",`"SPUserID`":`"$email`",`"AccountName`":`"$email`",`"PrincipalType`":`"UNVALIDATED_EMAIL_ADDRESS`"},`"MultipleMatches`":[],`"AutoFillKey`":`"$email`",`"AutoFillDisplayText`":`"$email`",`"AutoFillSubDisplayText`":`"`",`"AutoFillTitleText`":`"$email\n$email`",`"DomainText`":`"$myDomain`",`"Resolved`":true}]"
                    $sharingResult = [Microsoft.SharePoint.Client.Web]::ShareObject($clientContext, $Url, $peoplePickerValue, "group:$groupNumber", $groupNumber, $false, $false, $false, "", "")
                    $clientContext.Load($sharingResult)
                    $clientContext.ExecuteQuery()

                    Write-Host "Emailing user: " $email -ForegroundColor Green        
                    $invitationLink = $sharingResult.InvitedUsers[0].InvitationLink
                    $todaysDate = Get-Date -Format D
                    $emailBody = "<h3 style=`"color: red`">Test HTML email</h3><a href=`"$invitationLink`">Click this link to accept the invitation.</a>"
                    Send-MailMessage -To $email -From $username -Subject $emailSubject -Body $emailBody -BodyAsHtml -SmtpServer smtp.office365.com -UseSsl -Credential $psCredentials -Port 587
                }
                Write-Host "--- End of Batch ---"
                Start-Sleep -Seconds $BatchInterval
            }
        }
    }
}