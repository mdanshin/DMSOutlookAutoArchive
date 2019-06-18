# DMSOutlookAutoArchive
Move Outlook items from mailbox to archive, e.g. PST file

If PSSecurityException occur, try the following:
$cert = Get-ChildItem Cert:\LocalMachine\My\
Set-AuthenticodeSignature -Certificate $cert -FilePath "Path to file"
