<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>

##
function New-Config {
    [System.XML.XMLDocument]$XML = New-Object System.XML.XMLDocument
    
    [System.XML.XMLElement]$Root = $XML.CreateElement("config")
    $XML.appendChild($Root)

    [System.XML.XMLElement]$exchangeAccount = $Root.AppendChild($XML.CreateElement("exchangeAccount"))
    $exchangeAccount.InnerText = "username@domain.com"

    [System.XML.XMLElement]$pstFile = $Root.AppendChild($XML.CreateElement("pstFile"))
    $pstFile.InnerText = "Archive"

    [System.XML.XMLElement]$fromFolder = $Root.AppendChild($XML.CreateElement("fromFolder"))
    $fromFolder.InnerText = "Inbox"

    [System.XML.XMLElement]$toFolder = $Root.AppendChild($XML.CreateElement("toFolder"))
    $toFolder.InnerText = "Inbox Archive"

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
    
    [CmdletBinding()]
    param (
        [Parameter(Position=0,mandatory=$true)]
        $namespace
    )

    $namespace.Folders | Format-Table name
}

function Move-Items ($items, $archive) {
    $confirmation = Read-Host "Are you Sure You Want To Proceed [y/N]?"

    if ($confirmation -eq 'y' -and $items) {
        $deletedItems = $items | ForEach-Object -Process { $PSItem.Move($archive) }        
    }
    Write-Output ("Moved: " + ( $deletedItems | measure-object ).Count)
}

function New-Outlook {
    $outlook = New-Object -ComObject outlook.application
    $namespace = $outlook.Getnamespace("MAPI")
    return $namespace
}

function Read-Config {
    try {
        $config = [xml](Get-Content .\config.xml -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        Write-Error "config.xml does not exist. Try to use -NewConfig parametr."
        Break
    }

    [hashtable]$return = @{}

    $return.exchangeAccount = $config.config.exchangeAccount
    $return.pstFile = $config.config.pstFile
    $return.fromFolder  = $config.config.fromFolder
    $return.toFolder  = $config.config.toFolder    
    $return.moveDays = $config.config.moveDays
    $return.moveDate = $config.config.moveDate
    $return.oldest   = $config.config.oldest

    return $return

}
# SIG # Begin signature block
# MIIO+wYJKoZIhvcNAQcCoIIO7DCCDugCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCBBQ+4QN1BJk3qRjy4tblVtx
# OFegggxMMIIGGTCCBQGgAwIBAgIKPKVcrgAEAAKr2zANBgkqhkiG9w0BAQsFADBq
# MRIwEAYKCZImiZPyLGQBGRYCcnUxFzAVBgoJkiaJk/IsZAEZFgdpYnNjb3JwMRQw
# EgYKCZImiZPyLGQBGRYEcm9vdDETMBEGCgmSJomT8ixkARkWA2liczEQMA4GA1UE
# AxMHSUJTIENBMTAeFw0xOTAyMjcwNzI5NTdaFw0xOTEyMDYxMDE1MTJaMDsxCzAJ
# BgNVBAYTAlJVMQwwCgYDVQQKEwNJQlMxCzAJBgNVBAsTAklUMREwDwYDVQQDDAhJ
# QlNfY29kZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALZbzhF0Bbm8
# fNN5UF9zjumXMVfHygBN95BOnt3qwqEo9kyfkb2hfnnUWwTPtgU3ZLd5NVWgaimb
# dZT/k/dEYICHbBPJs0/xzqszzFwdY9Jqhl4erpsUYJHZKEei7fMqDhooVpW+8igP
# YOsREpUlj+bkY3yB2BFfBfWPjCww8KWYzqK6+gf91fUzjewE+uXoVTs1aaNjWHri
# tqj8YC29n2v3AvibAhgV9F/2DbAEFvXsAPXy2j5ZGWFNVKpjpThFktLVEk64Gdlg
# oliGxm2fpbfi9Zj+eXCu/9pG7ElBthbQvyQqwLS8KMwrA2Wew8aKOwFFC1yCL96d
# yGOYP2QXIrsCAwEAAaOCAu4wggLqMAsGA1UdDwQEAwIHgDA8BgkrBgEEAYI3FQcE
# LzAtBiUrBgEEAYI3FQiHmoMrgrD8MoPJlQuGvtIQo7kbYoHz3RiBrO1rAgFkAgEC
# MB0GA1UdDgQWBBRrtog/p2PNcETcGIFgQyjv+mpUNzAfBgNVHSMEGDAWgBQb1Vne
# FzlSKrfZGwtACMfRVVHy9DCCAQUGA1UdHwSB/TCB+jCB96CB9KCB8YaBvmxkYXA6
# Ly8vQ049SUJTJTIwQ0ExKDQpLENOPWhxLWliLWNhLTAxLENOPUNEUCxDTj1QdWJs
# aWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9u
# LERDPXJvb3QsREM9aWJzY29ycCxEQz1ydT9jZXJ0aWZpY2F0ZVJldm9jYXRpb25M
# aXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnSGLmh0dHA6
# Ly9jZXJ0Lmlicy5ydS9DZXJ0RW5yb2xsL0lCUyUyMENBMSg0KS5jcmwwggEgBggr
# BgEFBQcBAQSCARIwggEOMIGvBggrBgEFBQcwAoaBomxkYXA6Ly8vQ049SUJTJTIw
# Q0ExLENOPUFJQSxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNl
# cyxDTj1Db25maWd1cmF0aW9uLERDPXJvb3QsREM9aWJzY29ycCxEQz1ydT9jQUNl
# cnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0
# eTBaBggrBgEFBQcwAoZOaHR0cDovL2NlcnQuaWJzLnJ1L0NlcnRFbnJvbGwvaHEt
# aWItY2EtMDEuaWJzLnJvb3QuaWJzY29ycC5ydV9JQlMlMjBDQTEoNCkuY3J0MBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYBBQUHAwMw
# DQYJKoZIhvcNAQELBQADggEBAGXZy+QQvzd6dm9g41XsJdTY70hiSlgoitWF32+Y
# iPlMS8uvaQjXRF+pPEgZDboRL7P4pI+SutAjlejjptOAQCkuZsVttFjsBzSbUiVg
# 1ceNMt5Xd4IeDeOsaZBjyW8Q36wa44rhU0zZvQDM0UBWYuK4s0fkFkDNx/RT+ue4
# MzKJsGdjVAGDKAuR3JFlx149h1vzLwWMf0/sscUZ3kdjIClbFPKW9GsZhCx0RrrB
# LgWYRmr6hzTiJtsrMbB6Azx2ZcTgBv50ckUfFnyh0+YHlf6cjX1cmzw4egVBdqxJ
# wWbGHfnQY5iu1cfHvb7TjaxhWEAjBXNWbrFAGKbtBaKRTlIwggYrMIIEE6ADAgEC
# Agoa9BeOAAIAAAAUMA0GCSqGSIb3DQEBCwUAMBYxFDASBgNVBAMTC0lCUyBSb290
# IENBMB4XDTE1MTIwNjEwMDUxMloXDTE5MTIwNjEwMTUxMlowajESMBAGCgmSJomT
# 8ixkARkWAnJ1MRcwFQYKCZImiZPyLGQBGRYHaWJzY29ycDEUMBIGCgmSJomT8ixk
# ARkWBHJvb3QxEzARBgoJkiaJk/IsZAEZFgNpYnMxEDAOBgNVBAMTB0lCUyBDQTEw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC9ljTjunH2kcVA/a9sVxnR
# u+Nfhsw9bKgi1zm3/rxJcE3LlSx1vaI26aRiL/LHctISgELTV2EZco8wUJAV99Zd
# +OfzcviIQ0Kt4TIsxCEsSpGrcSF+UI0UFce2IEgLoMT9E9ZW5pYBiSejO3FB1Xqx
# qSTREOwWdyv2IKxGLsglv4agqRW+extFp/Wb66lNkXjq/n6JCiOIaBAjk/x2n9/2
# t9sw65LMbu4ryHOBywKgA0NRZXFfYmGccU1Ult/aT1X1MsYugseV4iTEwSMeBtbY
# fkOdx/E/NIgyZxta/8xz8Fwqz85vQwclVvV23N5z8fb56i4Z0fUu5fVP8t6a1Wrz
# AgMBAAGjggIlMIICITASBgkrBgEEAYI3FQEEBQIDBAAEMCMGCSsGAQQBgjcVAgQW
# BBQM70JXl1qPso90360iy0BBEVwf0jAdBgNVHQ4EFgQUG9VZ3hc5Uiq32RsLQAjH
# 0VVR8vQwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8G
# A1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU9M3ZNTLuMzlZZrXSblKpgb2gZoIw
# ggENBgNVHR8EggEEMIIBADCB/aCB+qCB94aBwWxkYXA6Ly8vQ049SUJTJTIwUm9v
# dCUyMENBLENOPWhxLWliLWNhLTAwLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBT
# ZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPXJvb3QsREM9
# aWJzY29ycCxEQz1ydT9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2Jq
# ZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnSGMWh0dHA6Ly9jZXJ0Lmlicy5y
# dS9DZXJ0RW5yb2xsL0lCUyUyMFJvb3QlMjBDQS5jcmwwXAYIKwYBBQUHAQEEUDBO
# MEwGCCsGAQUFBzAChkBodHRwOi8vY2VydC5pYnMucnUvQ2VydEVucm9sbC9ocS1p
# Yi1jYS0wMF9JQlMlMjBSb290JTIwQ0EoMikuY3J0MA0GCSqGSIb3DQEBCwUAA4IC
# AQBKOCQzVsAVEmjoIIPMxAtBRkoyxkToMjaZLxhPx5dCCKrnqsKMzpKNDR0Wsisr
# F0mr2LQvfDnGyUgdHFzZORPAOVNG0p759s0pXaMRPtgn/vmIDXR3+wFJrAsJ63YC
# LF+DbzWRe91suh5lbt/+VeSDIgexgz8vHfAkh1doVZk0yzOdnNdlJB0jKsJTAVZ9
# 0CTtCDXv54s4lNWQnXPuOcBcLZ3qUtW4+kMvmhqy2rUKSF7zdYI8uavnH+NAIso0
# kc2bd0zanlIxY5jgJoENO0J4XJ6M3ClgqnfYuHQJVNLQ0EYjSSZS4Ldt/ncI+ubB
# E6qqJBX0YwAYzD+dOZQDO2weQXa1Rb0CYiAQ912ZNbc5Z/VjjqEvtfCDAeGbkTOr
# GvohUzYnqLllCiqG5tNkJWRdiA65ztDzA4Ejws85iDwkSMFORMQrTTN4fDnzW5X9
# UheZGojX2au3QMGFU6tH0UHqkiY7lnxkYl/F0eajXizHDgQwaM5h5CoTsgBfAB/m
# A19uvP/5Pg58qazXS3d7RTvk3+ISiaO0lJgHVyrCXzncsrHVkXSIB1ibx+BByyLw
# xMhUMU7oqutvsWVUBInCqIebVZu7DI7xA0YwU1abtGvcvZmeIrBm1M6bE08DZCth
# MzcW6nUh6B6mOWUhWL4X+MrtcceqKyW5rmILGgWuXyS/8zGCAhkwggIVAgEBMHgw
# ajESMBAGCgmSJomT8ixkARkWAnJ1MRcwFQYKCZImiZPyLGQBGRYHaWJzY29ycDEU
# MBIGCgmSJomT8ixkARkWBHJvb3QxEzARBgoJkiaJk/IsZAEZFgNpYnMxEDAOBgNV
# BAMTB0lCUyBDQTECCjylXK4ABAACq9swCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcC
# AQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYB
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFLHUGNeeFCvU
# D5bsae/PSBR7RiSaMA0GCSqGSIb3DQEBAQUABIIBABw2G7RZYP7vvUFZC1PTV6Js
# sbwl/+q/MgOV9CHDiZP5IfBemwEPJ9PbvSZ/oGNfntv1GHZ1v8UX2laVydMrV02C
# Np1U/vk2wl61lhoMbRkhp/oHZ/J7CaBEPLNybjLpgJCOKB4H+jl8u/6KdR71WBMj
# n/YO2ltTjpC8jTssyXp8Bxg0tYSCI86puVXkctcLf8Q7lrSHcSTLutTbvozwyyuJ
# ZkIcbH4IqCFRTBb8x4zUGnIHuZL03nvkoC89u6qhCYpcPAX2wkd4GCUZ0ahNxkzn
# Zq/e56B4K5JLxxbjel/2BhDp/rx13tEOfkMHXQ/6h6aFkzDjxe9tXoNL1KT4bgo=
# SIG # End signature block
