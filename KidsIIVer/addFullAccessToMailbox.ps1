Write-Host "Getting O365 credentials" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

$credentials = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Write-Host "Enter mailbox to give access to" -ForegroundColor Green -BackgroundColor Black
$mToAccess = Read-Host

Write-Host "Enter user that needs access" -ForegroundColor Green -BackgroundColor Black
$oUser = Read-Host

Write-Host "Adding Full Access to" $mToAccess -ForegroundColor Green -BackgroundColor Black
Add-MailboxPermission -Identity $mToAccess -User $oUser -AccessRights FullAccess -InheritanceType All -Confirm:$false 

Start-Sleep -Seconds 1

$o365Mailbox = Get-MailboxPermission $mToAccess | where {($_.User -like $oUser)} | select -ExpandProperty User
if ($oUser -eq $o365Mailbox)
{
    Write-Host "Full Access Found" -ForegroundColor Green -BackgroundColor Black
    Start-Sleep -Seconds 5
}
else
{
    Write-Host "Failed To Add Full Access"
    Start-Sleep -Seconds 2
}



