Param(
    [Parameter(ParameterSetName='One')][switch]$Accounts,
    [Parameter(ParameterSetName='Two')][switch]$WhatIf,
    [Parameter(ParameterSetName='Three')][switch]$NewConfig
)

. ".\lib.ps1"

if ($Accounts) {
    Get-Accounts
    Exit
}

if ($NewConfig) {
    if (![System.IO.File]::Exists("$PWD\config.xml")) {
        New-Config
    }
    Exit
}

try {
    $config = [xml](Get-Content .\config.xml -ErrorAction Stop)
}
catch {
    Write-Error "config.xml does not exist. Try to use -NewConfig parametr."
    Break
}

$mAccount = $config.config.mainAccount
$aAccount = $config.config.archiveAccount
$moveDays = $config.config.moveDays
$moveDate = $config.config.moveDate

if ($config.config.oldest -eq 'true') {
    $oldest = $true
}

if ($config.config.oldest -eq 'false') {
    $oldest = $false
}

if ($moveDate) {
    [DateTime]$Date = $moveDate
}
else {
    $Date = [DateTime]::Now.AddDays(-$moveDays) 
}

$LastMonths =  $Date.tostring("MM/dd/yyyy")

$outlook = New-Object -ComObject outlook.application
$namespace = $outlook.Getnamespace("MAPI")

$mainAccount = $namespace.Folders | Where-Object { $_.Name -eq $mAccount };
$archiveAccount = $namespace.Folders | Where-Object { $_.Name -eq $aAccount };

$inbox = $mainAccount.Folders | Where-Object { $_.Name -match 'Sent Items'}
$archive = $archiveAccount.Folders | Where-Object { $_.Name -match 'Sent Items'}

Write-Output ("Total items: " + ($inboxItems = $inbox.Items).Count)

if ($oldest) {
    Write-Output ("Older then $LastMonths" + ": " + ($items = $inboxItems | Where-Object -FilterScript { $_.senton -le $LastMonths}).Count)
}
else {
    Write-Output ("Younger then $LastMonths" + ": " + ($items = $inboxItems | Where-Object -FilterScript { $_.senton -ge $LastMonths}).Count)
}

Move-Items