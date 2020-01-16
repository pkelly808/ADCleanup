function Get-PKAdcUser {

<#
.SYNOPSIS
ADCleanup for ADComputers.  Get-PKAdcUser will only retrieve data and specify what Action to take but no changes are made.
.DESCRIPTION
If Computer is Enabled and LastLogonDate is older than DisableDays, the Action is Disable.
If Computer is Disabled and older than RemoveDays, the Action is Remove.
If the OS is like *Server* or Unknown, it is skipped.
To exclude the computer and set the Action to Keep.  Add Keep to the computer description.

Input Parameters require positive integers

Output Object with the following Properties:
    ComputerName
    Enabled
    Action - Disable Wait Remove None New Keep
    LastLogonDate
    OperatingSystem
    Description

.EXAMPLE
Get-PKAdcUser l-w7-pkelly | ? {($_.Action -ne 'None')}
Gets the computers and displays Action.  None is taken.  Display Objects with Action not equal to None
.EXAMPLE
Send-PKAdcUser
A report is emailed but no Action is taken
.EXAMPLE
Get-ADComputer -Filter * | Get-PKAdcUser | Remove-PKAdcUser
Gets the computers and applies the Action.  No Output.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory,ValueFromPipeline)]
    [string[]]$Identity,

    [ValidatePattern('^[1-9]')]
    [int]$DisableDays = 90,

    [ValidatePattern('^[1-9]')]
    [int]$RemoveDays = 180
)

    BEGIN {
        $DisableDate = (Get-Date).AddDays(-$DisableDays)
        $RemoveDate = (Get-Date).AddDays(-$RemoveDays)

        #Combines days and used if computer is manually disabled without a date in the description
        $NoDescRemoveDate = (Get-Date).AddDays(-$DisableDays-$RemoveDays)
    }

    PROCESS {
 
        foreach ($User in $Identity) {

            try {
                $ADUser = Get-ADUser $User -Properties LastLogonDate,Description,whenCreated -ea Stop
            } catch {
                Write-Warning $_
                continue
            }

            if ($ADUser -ne $null) {

                if ($ADUser.LastLogonDate -lt $DisableDate) {
                    if ($ADUser.Enabled) {
                        Write-Verbose "$($ADUser.SamAccountName) Enabled, LastLogonDate older than DisableDate"
                        $Action = 'Disable'
                    } else {
                        try {
                            $DescDate = [datetime]$ADUser.Description.Trim("INACTIVE ").Split()[0]

                            if ($DescDate -lt $RemoveDate) {
                                Write-Verbose "$($ADUser.SamAccountName) Disabled, DescriptionDate older than RemoveDate"
                                $Action = 'Remove'
                            } else {
                                Write-Verbose "$($ADUser.SamAccountName) Disabled, DescriptionDate not old enough to Remove yet"
                                $Action = 'Wait'
                            }
                        } catch {
                            if ($ADUser.LastLogonDate -lt $NoDescRemoveDate) {
                                Write-Verbose "$($ADUser.SamAccountName) Disabled, No DescriptionDate, LastLogonDate older than combined Disable Remove Date"
                                $Action = 'Remove'
                            } else {
                                Write-Verbose "$($ADUser.SamAccountName) Disabled, No DescriptionDate, LastLogonDate not old enough to Remove yet from combined Disable Remove Date"
                                $Action = 'Wait'
                            }
                        }
                    }
                } else {
                        Write-Verbose "$($ADUser.SamAccountName) Active, LastLogonDate not older than DisableDate"
                        $Action = 'None'
                }

                if ($ADUser.SamAccountName -like 'svc*') {
                    Write-Verbose "$($ADUser.SamAccountName) Svc Account, Overrides all Actions"
                    $Action = 'Svc'
                }

                #Protects new users without LastLogonDate
                if ($ADUser.whenCreated -gt $DisableDate) {
                    Write-Verbose "$($ADUser.SamAccountName) New Account, Overrides all Actions"
                    $Action = 'New'
                }

                if ($ADUser.Description -like '*KEEP*') {
                    Write-Verbose "$($ADUser.SamAccountName) Description Keep, Overrides all Actions"
                    $Action = 'Keep'
                }

                [PSCustomObject]@{
                    SamAccountName=$ADUser.SamAccountName
                    Enabled=$ADUser.Enabled
                    Action=$Action
                    LastLogonDate=$ADUser.LastLogonDate
                    Description=$ADUser.Description
                }
            }
        }
    }
}

function Remove-PKAdcUser {

[CmdletBinding()]
param(
    [Parameter(Mandatory,ValueFromPipeline)]
    [psobject]$InputObject,

    [string]$ArchivePath = '\\fileshare\ADCleanup\PKAdcUser\'
)

    BEGIN {
        $Description = "INACTIVE $(Get-Date -f d)"

        if (!(Test-Path $ArchivePath)) {
            Write-Warning "$ArchivePath does not exist!"
            Exit
        }
    }

    PROCESS {

        foreach ($User in $InputObject) {

            Switch ($User.Action) {
                Disable {
                    Write-Verbose "$($User.SamAccountName) processing Disable: Remove ProtectedFromAccidentalDeletion, Disable, Update Description"

                    Get-ADUser $User.SamAccountName | Set-ADObject -ProtectedFromAccidentalDeletion $false -Confirm:$false

                    Set-ADUser $User.SamAccountName -Enabled $false -Description "$Description $($User.Description)" -Confirm:$false

                    }

                Remove {
                    Write-Verbose "$($User.SamAccountName) processing Remove: Remove-ADComputer"

                    Get-ADUser $User.SamAccountName -Properties * | Export-Clixml "$ArchivePath$($User.SamAccountName).xml"

                    #Remove-ADUser $User.SamAccountName -Confirm:$false

                    }

                Default {
                    Write-Verbose "$($User.SamAccountName) Nothing to process"
                    }
            }
        }
    }
}

function Send-PKAdcUser {

[CmdletBinding()]
param(
    [string]$SmtpServer = 'smtp-relay.gmail.com',
    [string]$From = 'Email <user@domain.com>',
    [string[]]$To = @('user@domain.com','user@domain.com'),
    [string]$Subject = 'ADCleanup Report: Users',

    [Parameter(Mandatory,ValueFromPipeline)]
    [psobject]$DS
)

    BEGIN {if (!(Get-Module EnhancedHTML2)) {Import-Module EnhancedHTML2}}

    PROCESS {

        $style = @"
        <style>
        body {
            color:#333333;
            font-family:Calibri,Tahoma;
            font-size: 10pt;
        }
        h1 {
            text-align:center;
        }
        h2 {
            border-top:1px solid #666666;
        }
        hr {
            height:1px;
            border-width:0;
            background-color:gray; }

        th {
            font-weight:bold;
            color:#eeeeee;
            background-color:#333333;
            padding:4px;
        }
        .odd  { background-color:#ffffff; }
        .even { background-color:#dddddd; }

        .dataTables_info { margin-bottom:4px; }
        .grid { width:100% }
        .red   { text-align:right; font-weight:bold; color:red; }
        .bold  { text-align:right; font-weight:bold; }
        .right { text-align:right; }
        </style>
"@

        $params = @{'As'='Table';
                    'PreContent'="<h2>&diams; ADCleanup Users</h2>The 'Action' has been applied to the following users.<br>Add KEEP in the Description to exclude from ADCleanup.<br>Default: Disable inactive after 90 days, Remove disabled after 180 days."
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                }
        $html_PKAdcUser = $ds | ConvertTo-EnhancedHTMLFragment @params

        $params = @{'CssStyleSheet'=$style;
                    'Title'=$Subject;
                    'PreContent'="<h1>$Subject</h1>";
                    'PostContent'="<br><hr /><i>Created $(Get-Date)</i>";
                    'HTMLFragments'=@($html_PKAdcUser)} 
        $Body = ConvertTo-EnhancedHTML @params

        Send-MailMessage -SmtpServer $SmtpServer -From $From -To $To -Subject $Subject -Body ($Body | Out-String) -BodyAsHtml
    
    }
}


function Use-PKAdcUser {

    $DS = Get-ADUser -Filter * -ResultSetSize 10000 | Get-PKAdcUser | ? Action -eq 'Disable' | sort Action,SamAccountName

    if ($DS) {
        Send-PKAdcUser -DS $DS

        #$DS | Remove-PKAdcUser
    }    
}

function Register-PKAdcUser {
    $Cred = Get-Credential -Message 'Account used to run Scheduled Task.'

    if ($Cred) {
        $params = @{
            Action=New-ScheduledTaskAction -Execute 'powershell.exe' -Argument 'Use-PKAdcUser'
            Trigger=New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tu -At 10am
            User=$Cred.UserName
            Password=$Cred.GetNetworkCredential().Password
            TaskName='PKAdcUser'
            Description='Disable/Remove stale AD User Objects and email report. -PKADCleanup'
            RunLevel='Highest'
        }

        Register-ScheduledTask @params -Force
    }
}