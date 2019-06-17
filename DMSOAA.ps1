function New-Config {
    [System.XML.XMLDocument]$XML = New-Object System.XML.XMLDocument
    
    [System.XML.XMLElement]$Root = $XML.CreateElement("config")
    $XML.appendChild($Root)

    [System.XML.XMLElement]$mainAccount = $Root.AppendChild($XML.CreateElement("mainAccount"))
    $mainAccount.InnerText = "username@domain.com"

    [System.XML.XMLElement]$archiveAccount = $Root.AppendChild($XML.CreateElement("archiveAccount"))
    $archiveAccount.InnerText = "Archive"

    [System.XML.XMLElement]$moveDays = $Root.AppendChild($XML.CreateElement("moveDays"))
    $comment = $XML.CreateComment('Not used if moveDate is set')
    $XML.DocumentElement.AppendChild($comment)
    $moveDays.InnerText = "30"

    [System.XML.XMLElement]$moveDate = $Root.AppendChild($XML.CreateElement("moveDate"))
    $comment = $XML.CreateComment('MM/dd/yyyy')
    $XML.DocumentElement.AppendChild($comment)
    $moveDate.InnerText = ""

    [System.XML.XMLElement]$oldest = $Root.AppendChild($XML.CreateElement("oldest"))
    $oldest.InnerText = "true"

    $XML.Save(("$pwd\config.xml"))
}

function get-accounts {
    $namespace.Folders | Format-Table name
}

if (![System.IO.File]::Exists("$PWD\config.xml")) {
    New-Config
}

$config = [xml](Get-Content .\config.xml)

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

$deletedItems = $items | ForEach-Object -Process { $PSItem.Move($archive) }
Write-Output ("Moved: " + $deletedItems.Count)