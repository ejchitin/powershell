Write-Host "Getting O365 credentials" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

$credentials = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Write-Host "Enter mailbox to remove access from" -ForegroundColor Green -BackgroundColor Black
$mToAccess = Read-Host

Write-Host "Enter user being removed" -ForegroundColor Green -BackgroundColor Black
$oUser = Read-Host

Write-Host "Removing Full Access from" $mToAccess -ForegroundColor Green -BackgroundColor Black
Remove-MailboxPermission -Identity $mToAccess -User $oUser -AccessRights FullAccess -InheritanceType All -Confirm:$false 

Start-Sleep -Seconds 1

$o365Mailbox = Get-MailboxPermission $mToAccess | where {($_.User -like $oUser)} | select -ExpandProperty User
if ($oUser -eq $o365Mailbox)
{
    Write-Host "Failed to remove Full Access"-ForegroundColor Green -BackgroundColor Black
    Start-Sleep -Seconds 5
}
else
{
    Write-Host "Full Access removed"
    Start-Sleep -Seconds 2
}



