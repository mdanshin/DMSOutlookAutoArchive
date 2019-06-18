# DMSOutlookAutoArchive
Move Outlook items from mailbox to archive, e.g. PST file

If PSSecurityException occur, try the following:
<br>
$cert = Get-ChildItem Cert:\LocalMachine\My\
Set-AuthenticodeSignature -Certificate $cert -FilePath "Path to file"

#How to use
TBD

#TODO
1. Update New-Config function
2. Update config.xml.example
3. Add user confirmation
