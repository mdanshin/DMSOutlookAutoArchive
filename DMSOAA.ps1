[CmdLetBinding(DefaultParameterSetName="None")]
Param(
    [Parameter(ParameterSetName='Accounts', Mandatory=$true)]
    [switch]$Accounts,

    [Parameter(ParameterSetName='WhatIf', Mandatory=$false)]
    [switch]$WhatIf,

    [Parameter(ParameterSetName='NewConfig', Mandatory=$false)]
    [switch]$NewConfig
)

. ".\lib.ps1"

switch ($PsCmdlet.ParameterSetName) {
    "Accounts" {
        Get-Accounts (New-Outlook)
        Exit
    }
    "NewConfig" {
        if (![System.IO.File]::Exists("$PWD\config.xml")) {
        New-Config
        }
    }
    "WhatIf" {
        Write-Host "TBD"

    }
    Default {

        try {
            $config = [xml](Get-Content .\config.xml -ErrorAction Stop)
        }
        catch {
            Write-Error "config.xml does not exist. Try to use -NewConfig parametr."
            Break
        }

        $config = Read-Config

        if ($config.moveDate) {
            [DateTime]$Date = $config.moveDate
        }
        else {
            $Date = [DateTime]::Now.AddDays(-$config.moveDays) 
        }

        $deleteDate =  $Date.tostring("MM/dd/yyyy")

        $outlook = New-Object -ComObject outlook.application
        $namespace = $outlook.Getnamespace("MAPI")

        $mainAccount = $namespace.Folders | Where-Object { $_.Name -eq $config.mAccount };
        $archiveAccount = $namespace.Folders | Where-Object { $_.Name -eq $config.aAccount };

        $inbox = $mainAccount.Folders | Where-Object { $_.Name -match 'Sent Items'}
        $archive = $archiveAccount.Folders | Where-Object { $_.Name -match 'Sent Items'}

        Write-Output ("Total items: " + ($inboxItems = $inbox.Items).Count)

        switch ($config.oldest) {
            'true' {
                Write-Output ("Older then $deleteDate" + ": " + ( $items = $inboxItems | Where-Object -FilterScript { $_.senton -le $deleteDate}).Count)
            }
            Default {
                Write-Output ("Younger then $deleteDate" + ": " + ($items = $inboxItems | Where-Object -FilterScript { $_.senton -ge $deleteDate}).Count)
            }
        }

        Move-Items $items $archive
    }
}
