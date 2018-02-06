# Powershell_Send-WefflesEmail
Powershell module for parsing weffles csv logs and sending emails based on results.

# Weffles
https://blogs.technet.microsoft.com/jepayne/2017/12/08/weffles/

# Examples
Send-WefflesEmail -mode summary -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -priorDays 1

Send-WefflesEmail -mode summary -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -startDate 1/1/2018 -endDate 12/31/2018

Send-WefflesEmail -mode alert -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -priorDays 1

Send-WefflesEmail -mode alert -emailTo admin@example.com -emailFrom alerts@example.com -smtpServerAddress smtp.office365.com -smtoServerPort 587 -senderPassword password -wefflesCsv C:\weffles\weffles.csv -startDate 1/1/2018 -endDate 12/31/2018

# Notes
If you don't want to pass email and csv path info each time the defaults can be changed in the params section of the script.
Example:

param (
    
    [ValidateSet('summary','alert')]
    [string]$mode = 'summary',
    [string]$emailTo = 'admin@example.com',
    [string]$emailFrom = 'alerts@example.com',
    [string]$smtpServerAddress = 'smtp.office365.com',
    [int]$smtoServerPort = 587,
    [string]$senderPassword = 'password',
    [string]$wefflesCsv = 'C:\weffles\weffles.csv',
    [int[]]$urgentEventIds = @(1102,7045,4720),
    [parameter(ParameterSetName="history")]
    [int]$priorDays = 1,
    [parameter(ParameterSetName="range")]
    [DateTime]$startDate = '1/1/2018',
    [parameter(ParameterSetName="range")]
    [DateTime]$endDate = '12/31/2018'
)

Then the script could be run like this:
Send-WefflesEmail -mode summary -priorDays 1

Send-WefflesEmail -mode summary -startDate 1/1/2018 -endDate 12/31/2018

Send-WefflesEmail -mode alert -priorDays 1

Send-WefflesEmail -mode alert -startDate 1/1/2018 -endDate 12/31/2018
