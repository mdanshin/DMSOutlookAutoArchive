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

function Get-Accounts {
    $namespace.Folders | Format-Table name
}

function Move-Items {
    if ($WhatIf) {
        
    }
    else {
        $deletedItems = $items | ForEach-Object -Process { $PSItem.Move($archive) }
        Write-Output ("Moved: " + $deletedItems.Count)        
    }
}