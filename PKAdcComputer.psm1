function Get-PKAdcComputer {

<#
.SYNOPSIS
ADCleanup for ADComputers.  Get-PKAdcComputer will only retrieve data and specify what Action to take but no changes are made.
.DESCRIPTION
If Computer is Enabled and LastLogonDate is older than DisableDays, the Action is Disable.
If Computer is Disabled and older than RemoveDays, the Action is Remove.
If the OS is like *Server* or Unknown, it is skipped.
To exclude the computer and set the Action to Keep.  Add Keep to the computer description.

Input Parameters require positive integers

Output Object with the following Properties:
    ComputerName
    Enabled
    Action - Disable Wait Remove None Keep
    LastLogonDate
    OperatingSystem
    Description

.EXAMPLE
Get-PKAdcComputer <computer> | ? {($_.Action -ne 'None')}
Gets the computers and displays Action.  None is taken.  Display Objects with Action not equal to None
.EXAMPLE
Send-PKAdcComputer
A report is emailed but no Action is taken
.EXAMPLE
Get-ADComputer -Filter * | Get-PKAdcComputer | Remove-PKAdcComputer
Gets the computers and applies the Action.  No Output.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory,ValueFromPipeline)]
    [string[]]$ComputerName,

    [ValidatePattern('^[1-9]')]
    [int]$DisableDays = 30,

    [ValidatePattern('^[1-9]')]
    [int]$RemoveDays = 30
)

    BEGIN {
        $DisableDate = (Get-Date).AddDays(-$DisableDays)
        $RemoveDate = (Get-Date).AddDays(-$RemoveDays)

        #Combines days and used if computer is manually disabled without a date in the description
        $NoDescRemoveDate = (Get-Date).AddDays(-$DisableDays-$RemoveDays)
    }

    PROCESS {
 
        foreach ($Computer in $ComputerName) {

            try {
                $Comp = Get-ADComputer $Computer -Properties LastLogonDate,Description,OperatingSystem -ea Stop
            } catch {
                Write-Warning $_
                continue
            }

            #Logic to process servers and unknown OS can be added in the future
            if ($Comp.OperatingSystem -like '*server*' -or $Comp.OperatingSystem -eq $null) {
                Write-Verbose "$Computer is a server or unknown, continue to next computer"
                continue
            }

            if ($Comp -ne $null) {

                if ($Comp.LastLogonDate -lt $DisableDate) {
                    if ($Comp.Enabled) {
                        Write-Verbose "$($Comp.Name) Enabled, LastLogonDate older than DisableDate"
                        $Action = 'Disable'
                    } else {
                        try {
                            $DescDate = [datetime]$Comp.Description.Trim("INACTIVE ").Split()[0]

                            if ($DescDate -lt $RemoveDate) {
                                Write-Verbose "$($Comp.Name) Disabled, DescriptionDate older than RemoveDate"
                                $Action = 'Remove'
                            } else {
                                Write-Verbose "$($Comp.Name) Disabled, DescriptionDate not old enough to Remove yet"
                                $Action = 'Wait'
                            }
                        } catch {
                            if ($Comp.LastLogonDate -lt $NoDescRemoveDate) {
                                Write-Verbose "$($Comp.Name) Disabled, No DescriptionDate, LastLogonDate older than combined Disable Remove Date"
                                $Action = 'Remove'
                            } else {
                                Write-Verbose "$($Comp.Name) Disabled, No DescriptionDate, LastLogonDate not old enough to Remove yet from combined Disable Remove Date"
                                $Action = 'Wait'
                            }
                        }
                    }
                } else {
                        Write-Verbose "$($Comp.Name) Active, LastLogonDate not older than DisableDate"
                        $Action = 'None'
                }

                if ($Comp.Description -like '*KEEP*') {
                    Write-Verbose "$($Comp.Name) Description Keep, Overrides all Actions"
                    $Action = 'Keep'
                }
            
                [PSCustomObject]@{
                    ComputerName=$Comp.Name
                    Enabled=$Comp.Enabled
                    Action=$Action
                    LastLogonDate=$Comp.LastLogonDate
                    OperatingSystem=$Comp.OperatingSystem
                    Description=$Comp.Description
                }
            }
        }
    }
}

function Remove-PKAdcComputer {

[CmdletBinding()]
param(
    [Parameter(Mandatory,ValueFromPipeline)]
    [psobject]$InputObject
)

    BEGIN {
        $Description = "INACTIVE $(Get-Date -f d)"
    }

    PROCESS {

        foreach ($Computer in $InputObject) {

            Switch ($Computer.Action) {
                Disable {
                    Write-Verbose "$($Computer.ComputerName) processing Disable: Remove ProtectedFromAccidentalDeletion, Disable, Update Description"

                    Get-ADComputer $Computer.ComputerName | Set-ADObject -ProtectedFromAccidentalDeletion $false -Confirm:$false

                    Set-ADComputer $Computer.ComputerName -Enabled $false -Description "$Description $($Computer.Description)" -Confirm:$false

                    }

                Remove {
                    Write-Verbose "$($Computer.ComputerName) processing Remove: Remove-ADComputer"

                    Remove-ADComputer $Computer.ComputerName -Confirm:$false

                    }

                Default {
                    Write-Verbose "$($Computer.ComputerName) Nothing to process"
                    }
            }
        }
    }
}

function Send-PKAdcComputer {

[CmdletBinding()]
param(
    [string]$SmtpServer = 'smtp-relay.gmail.com',
    [string]$From = 'Email <user@domain.com>',
    [string[]]$To = @('user@domain.com','user@domain.com'),
    [string]$Subject = 'ADCleanup Report: Windows Computers'
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
                    'PreContent'="<h2>&diams; ADCleanup Computers</h2>The 'Action' has been applied to the following computers.<br>Add KEEP in the Description to exclude from ADCleanup.<br>Default: Disable inactive after 30 days, Remove disabled after 30 days."
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                }
        $html_PKAdcComputer = Get-ADComputer -Filter {OperatingSystem -like 'Windows*'} | Get-PKAdcComputer | Where {$_.Action -ne 'None'} | sort Action,ComputerName | ConvertTo-EnhancedHTMLFragment @params

        $params = @{'CssStyleSheet'=$style;
                    'Title'=$Subject;
                    'PreContent'="<h1>$Subject</h1>";
                    'PostContent'="<br><hr /><i>Created $(Get-Date)</i>";
                    'HTMLFragments'=@($html_PKAdcComputer)} 
        $Body = ConvertTo-EnhancedHTML @params

        Send-MailMessage -SmtpServer $SmtpServer -From $From -To $To -Subject $Subject -Body ($Body | Out-String) -BodyAsHtml
    }
}


function Use-PKAdcComputer {

    Send-PKAdcComputer 
    
    Get-ADComputer -Filter {OperatingSystem -like 'Windows*'} | Get-PKAdcComputer | Remove-PKAdcComputer
}

function Register-PKAdcComputer {
    $Cred = Get-Credential -Message 'Account used to run Scheduled Task.'

    if ($Cred) {
        $params = @{
            Action=New-ScheduledTaskAction -Execute 'powershell.exe' -Argument 'Use-PKAdcComputer'
            Trigger=New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tu -At 10am
            User=$Cred.UserName
            Password=$Cred.GetNetworkCredential().Password
            TaskName='PKAdcComputer'
            Description='Disable/Remove stale AD Computer Objects and email report.  Windows Only. -PKADCleanup'
            RunLevel='Highest'
        }

        Register-ScheduledTask @params -Force
    }
}