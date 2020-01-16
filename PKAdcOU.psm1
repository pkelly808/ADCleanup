function Measure-PKAdcOU {
[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline)]
    [string[]]$OU
)
    BEGIN {
        if (!$OU) {
            $OU = Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel
        }
    }
    
    PROCESS {
        foreach ($OUnit in $OU) {
            $User = Get-ADUser -Filter * -SearchBase $OUnit 

            $Comp = Get-ADComputer -Filter * -SearchBase $OUnit

            [PSCustomObject]@{
                OU=$OUnit
                UserEnabled=($User | Where-Object {$_.Enabled -eq $True} | Measure-Object).Count
                UserDisabled=($User | Where-Object {$_.Enabled -eq $false} | Measure-Object).Count
                UserTotal=($User | Measure-Object).Count
                CompEnabled=($Comp | Where-Object {$_.Enabled -eq $True} | Measure-Object).Count
                CompDisabled=($Comp | Where-Object {$_.Enabled -eq $false} | Measure-Object).Count
                CompTotal=($Comp | Measure-Object).Count
            }
        }
    }
}

function Send-PKAdcOU {

[CmdletBinding()]
param(
    [string]$SmtpServer = 'smtp-relay.gmail.com',
    [string]$From = 'Email <user@domain.com>',
    [string[]]$To = @('user@domain.com','user@domain.com'),
    [string]$Subject = 'ADCleanup Report: Users and Computers',

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
                    'PreContent'='<h2>&diams; Users and Computers</h2>'
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'Properties'='OU',
                        @{n='User<br>Enabled';e={$_.UserEnabled};css={ 'right' }},
                        @{n='User<br>Disabled';e={$_.UserDisabled};css={if ($_.UserDisabled -ne 0) { 'red' } else { 'right' }}},
                        @{n='User<br>Total';e={$_.UserTotal};css={ 'bold' }},
                        @{n='Computer<br>Enabled';e={$_.CompEnabled};css={ 'right' }},
                        @{n='Computer<br>Disabled';e={$_.CompDisabled};css={if ($_.CompDisabled -ne 0) { 'red' } else { 'right' }}},
                        @{n='Computer<br>Total';e={$_.CompTotal};css={ 'bold' }}
                        }
        $html_uc = $DS | ConvertTo-EnhancedHTMLFragment @params

        $params = @{'CssStyleSheet'=$style;
                    'Title'=$Subject;
                    'PreContent'="<h1>$Subject</h1>";
                    'PostContent'="<br><hr /><i>Created $(Get-Date)</i>";
                    'HTMLFragments'=@($html_uc)} 
        $Body = ConvertTo-EnhancedHTML @params

        Send-MailMessage -SmtpServer $SmtpServer -From $From -To $To -Subject $Subject -Body ($Body | Out-String) -BodyAsHtml
    }
}


function Use-PKAdcOU {

    $DS = Measure-PKAdcOU

    Send-PKAdcOU -DS $DS
}

function Register-PKAdcOU {
    $Cred = Get-Credential -Message 'Account used to run Scheduled Task.'

    if ($Cred) {
        $params = @{
            Action=New-ScheduledTaskAction -Execute 'powershell.exe' -Argument 'Use-PKAdcOU'
            Trigger=New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tu -At 10am
            User=$Cred.UserName
            Password=$Cred.GetNetworkCredential().Password
            TaskName='PKAdcOU'
            Description='Report OU with Enabled/Disabled User and Computer objects. -PKADCleanup'
            RunLevel='Highest'
        }

        Register-ScheduledTask @params -Force
    }
}