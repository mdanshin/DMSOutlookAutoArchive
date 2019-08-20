Param
      (
         [parameter(Mandatory=$true, ValueFromPipeline=$true)]
         $Hash
      )

if ($Hash -is [System.Collections.Hashtable])
{
    $dictionary = [ordered]@{}
    $keys = $Hash.keys | sort

    foreach ($key in $keys)
    {
        $dictionary.add($key, $Hash[$key])
    }

    return $dictionary
}
elseif ($Hash -is [System.Array])
{
    $dictionary = [ordered]@{}
    
    for ($i = 0; $i -lt $hash.count; $i++)
    {
        $dictionary.add($i, $hash[$i])
    }

    return $dictionary
}
else
{
    Write-Error "Enter a hash table or an array."
}
<#    
    .SYNOPSIS
    Converts a hash table or an array to an ordered dictionary. 
            
    .DESCRIPTION
    ConvertTo-OrderedDictionary takes a hash table or an array and 
    returns an ordered dictionary. 
    
    If you enter a hash table, the keys in the hash table are ordered 
    alphanumerically in the dictionary. If you enter an array, the keys 
    are integers 0 - n.
            
    .PARAMETER  $hash
    Specifies a hash table or an array. Enter the hash table or array, 
    or enter a variable that contains a hash table or array.

    .INPUTS
    System.Collections.Hashtable
    System.Array

    .OUTPUTS
    System.Collections.Specialized.OrderedDictionary

    .EXAMPLE
    PS C:\> $myHash = @{a=1; b=2; c=3}
    PS C:\> .\ConvertTo-OrderedDictionary.ps1 -Hash $myHash

    Name                           Value                                                                                                                                                           
    ----                           -----                                                                                                                                                           
    a                              1                                                                                                                                                               
    b                              2                                                                                                                                                               
    c                              3                          

    .EXAMPLE
    PS C:\> $myHash = @{a=1; b=2; c=3}
    PS C:\> $myHash = .\ConvertTo-OrderedDictionary.ps1 -Hash $myHash
    PS C:\> $myHash

    Name                           Value                                                                                                                                                           
    ----                           -----                                                                                                                                                           
    a                              1                                                                                                                                                               
    b                              2                                                                                                                                                               
    c                              3
                  

    PS C:\> $myHash | Get-Member
    
       TypeName: System.Collections.Specialized.OrderedDictionary
       . . .

    .EXAMPLE
    PS C:\> $colors = "red", "green", "blue"
    PS C:\> $colors = .\ConvertTo-OrderedDictionary.ps1 -Hash $colors
    PS C:\> $colors

    Name                           Value                                                                                                                                                           
    ----                           -----                                                                                                                                                           
    0                              red                                                                                                                                                             
    1                              green                                                                                                                                                           
    2                              blue 

 
    .LINK
    about_hash_tables
#>

# SIG # Begin signature block
# MIIO+wYJKoZIhvcNAQcCoIIO7DCCDugCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU54sUP75bz+e1VB06TLHGSoyP
# 21SgggxMMIIGGTCCBQGgAwIBAgIKPKVcrgAEAAKr2zANBgkqhkiG9w0BAQsFADBq
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFF1SfC/70VZJ
# 7b+Wp4jeXI2IiavKMA0GCSqGSIb3DQEBAQUABIIBAJ+j/IbtIjv35+o8JJ0txA58
# JYZ8K+PBJcZiC3TKu8DkUO8F7s6UeQ1s3HeJeVxyvSJYnJzxIcDb9R2neqRaVaDm
# U5KLDqUAJPvHEBwqQrbxSbMUPeSruf3qjuMh0jcgyLnsizaqYJsQ/EVj/9WL8SUn
# DbXFGCcpYngmeheUqKUMsNhd0kcZtmGgDHK7/+O8K2B1gLpgeLSr3vdgqCkOiUx3
# AhjCmlKNa/XQAvkfeyN+8kFT7czaCF/2xgs8wEFvv/XKGTFEitkW8HVNbnJDiwlR
# plyAQyX4FL1TE8kxrnN1388n/BgpAzJ48KBMkb+B0ZOU+5IYRiUzSFQv7tmQWEk=
# SIG # End signature block
