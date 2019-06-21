
# DMSOutlookAutoArchive

Move Outlook items from mailbox to archive, e.g. PST file

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software and how to install them

You will need PowerShell and Microsoft Outlook. Please note that Outlook must be properly configured to work with at least one mailbox and at least one PST file must be connected to it.

### How to use

Clone or download project and extract from archive.

First run the `DMSOAA.ps1` with `-NewConfig` parameter.

```Powershell
.\DMSOAA.ps1 -NewConfig
```

Then run the script with `-Accounts` parameter.

```Powershell
.\DMSOAA.ps1 -Accounts
```

You will see all connected mailboxes and data files.

Finally, edit the configuration file `config.xml`.Use the information you received before. Then run the script without any parameters.

## Example of the config.xml

```XML
<config>
  <exchangeAccount>username@domain.com</exchangeAccount> <!--A-->
  <pstFile>Archive</pstFile>
  <fromFolder>Sent Items</fromFolder>
  <toFolder>Sent Items</toFolder>
  <moveDays>30</moveDays> <!--Not used if moveDate is set-->
  <moveDate></moveDate> <!--MM/dd/yyyy-->
  <oldest>true</oldest>
</config>
```

## Authors

* **Mikhail Danshin** - *Initial work* - [mdanshin](https://github.com/mdanshin)

See also the list of [contributors](https://github.com/mdanshin/DMSOutlookAutoArchive/graphs/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

TBD

## Known issues
If PSSecurityException occur, try the following:

```Powershell
$cert = Get-ChildItem Cert:\LocalMachine\My\
Set-AuthenticodeSignature -Certificate $cert -FilePath "Path to file"
```

You may have problems working with Cyrillic characters.
