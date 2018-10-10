$creds = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $creds -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

new-mailbox -Name Room1 -Room
new-mailbox -Name Room2 -Room

Get-MailBox -RecipientTypeDetails roommailbox -Filter { Name -like "Room*" } | Set-CalendarProcessing -AutomateProcessing:AutoAccept

New-DistributionGroup -Name RoomList1 -RoomList
Add-DistributionGroupMember -Identity "RoomList1" -Member "Room1"
Add-DistributionGroupMember -Identity "RoomList1" -Member "Room2"


$cert=New-SelfSignedCertificate -Subject "CN=EWSResourceSync" -CertStoreLocation "Cert:\CurrentUser\My"  -KeyExportPolicy Exportable -KeySpec Signature
Export-PfxCertificate -Cert $cert -Password (Get-Credential).Password -FilePath .\temp.pfx -Verbose
Export-Certificate -Cert $cert -FilePath .\temp.cer -Type CERT


Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=EWSResourceSync" }
