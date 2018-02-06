<#
.SYNOPSIS
Send-WefflesEmail parses a weffles csv file and send emails based on rules.
.DESCRIPTION
Send-WefflesEmail parses a weffles csv file and send emails based on rules.
Alert mode will send an email notification if a urgent event id is found.
Summary mode will send an email summary of all events based over X days, or between two specific dates. 
.PARAMETER mode
The notification mode for emails. Default: summary.
.PARAMETER emailTo
The email address that will receive the notification emails
.PARAMETER emailFrom
The email address that will be used to send emails
.PARAMETER smtpServerAddress
The address of the server used for sending emails
.PARAMETER smtoServerPort
The port for smtp on the email server used for sending emails
.PARAMETER senderPassword
The password for the email account used to send emails
.PARAMETER wefflesCsv
The location of the weffles.csv file.  Default: C:\weffles\weffles.csv
.PARAMETER priorDays
For summary mode only.  How many days worth of logs to review.  Default: 1
.PARAMETER urgentEventIds
For alert mode only.  A list of event IDs that should cause an email alert.  Default: 1102,7045,4720
.EXAMPLE
Send-WefflesEmail -mode summary -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -priorDays 1
Send-WefflesEmail -mode summary -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -startDate 1/1/2018 -endDate 12/31/2018
Send-WefflesEmail -mode alert -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -priorDays 1
Send-WefflesEmail -mode alert -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -startDate 1/1/2018 -endDate 12/31/2018
#>

[CmdletBinding()]
param (
    [ValidateSet('summary','alert')]
    [string]$mode = 'summary',
    [string]$emailTo = '',
    [string]$emailFrom = '',
    [string]$smtpServerAddress = '',
    [int]$smtoServerPort = 587,
    [string]$senderPassword = '',
    [string]$wefflesCsv = 'C:\weffles\weffles.csv',
    [int[]]$urgentEventIds = @(1102,7045,4720),
    [parameter(ParameterSetName="history")]
    [int]$priorDays = 1,
    [parameter(ParameterSetName="range")]
    [DateTime]$startDate = '1/1/2018',
    [parameter(ParameterSetName="range")]
    [DateTime]$endDate = '12/31/2018'
)

#Functions
#SEND EMAIL
Function SendEmail {
    Param ($subject, $body)

    $credentialObject = New-Object System.Management.Automation.PSCredential ($emailFrom, (ConvertTo-SecureString $senderpassword -AsPlainText -Force))
    Send-MailMessage -To $emailTo -Body $body -From $emailFrom -Subject $subject -SmtpServer $smtpServerAddress -Port $smtoServerPort -UseSsl -Credential $credentialObject 
}

#Main
switch ($PsCmdlet.ParameterSetName) {
    "history”
    {
        $startDate = [DateTime]::Today.AddDays(-1 * $priorDays)
        $endDate = [DateTime]::Today
    } 
} 

#Values in CSV that we care about
$events = Import-Csv $wefflesCsv | foreach {
  New-Object PSObject -prop @{
    EventDate = [DateTime]::Parse($_.EventDate);
    EventHost = $_.EventHost
    EventID = [int]::Parse($_.EventID)
    #TargetUserName = $_.TargetUserName
  }
}

$events = $events | where-object {$_.EventDate -gt $startDate -and $_.EventDate -le $endDate} 
$groupedEvents = $events | Group-Object EventID

foreach ($eventId in $urgentEventIds){
    foreach ($group in $groupedEvents) {
        if ($group.Name -eq $eventId) {
            SendEmail -subject "Weffles Alarm" -body ("Event ID " + $eventId.ToString() + " detected by weffles " + $group.Count.ToString() +" between " + $startDate.ToShortDateString() + " and " + $endDate.ToShortDateString() + ".  This ID is defined as urgent and needs immediate attention.")          
            foreach ($instance in $group.Group) {               
                "Urgent! Date: " +$instance.EventDate.ToShortDateString() + " Host: " + $instance.EventHost + "  Event: " + $instance.EventID
            }
        }
    }
}

if ($mode -eq 'summary') {
    #make a summary
    $groupedEventsByDate = $events | Sort-Object EventDate,EventHost,EventID -unique | Group-Object EventDate

    $summary = "Date       EventID Count Host`n"
    foreach ($day in $groupedEventsByDate) {
        foreach ($event in $day.Group) {
            $uniqueCount = $events | where-object {$_.EventDate -eq $event.EventDate -and $_.EventID -eq $event.EventID -and $_.EventHost -eq $event.EventHost} | measure
            $summary += $event.EventDate.ToString("MM/dd/yyyy") + " " + $event.EventID.ToString() + "     " + $uniqueCount.Count + "    " + $event.EventHost.ToString() + "`n"
        }
    }

    SendEmail -subject "Weffles Summary" -body $summary
    $summary |ft
}