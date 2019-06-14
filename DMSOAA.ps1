$config = [xml](Get-Content .\config.xml)

$Date = [DateTime]::Now.AddDays(-30)
$LastMonths =  $Date.tostring("MM/dd/yyyy")

$mAccount = $config.Accounts.mainAccount
$aAccount = $config.Accounts.archiveAccount

$outlook = New-Object -ComObject outlook.application
$namespace = $outlook.Getnamespace("MAPI")

$mainAccount = $namespace.Folders | Where-Object { $_.Name -eq $mAccount };
$archiveAccount = $namespace.Folders | Where-Object { $_.Name -eq $aAccount };

$inbox = $mainAccount.Folders | Where-Object { $_.Name -match 'Sent Items'}
$archive = $archiveAccount.Folders | Where-Object { $_.Name -match 'Sent Items'}

$inbox.Items | Where-Object -FilterScript {
    $_.Sent -ge $LastMonths
}

