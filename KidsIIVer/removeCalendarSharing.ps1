Write-Host "Getting O365 Credentials" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2

$iCal = ":\Calendar"

$credentials = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Write-Host "Enter email of calendar to remove access from" -ForegroundColor Green -BackgroundColor Black
$calendarToShare = Read-Host

Write-Host "Enter user being removed" -ForegroundColor Green -BackgroundColor Black
$oUser = Read-Host

Write-Host "removing user..." -ForegroundColor Green -BackgroundColor Black
Remove-MailboxFolderPermission -Identity ($calendarToShare + ":\Calendar") -User $oUser 

Write-Host "Current permissions of: " $calendarToShare
Get-MailboxFolderPermission -Identity ($calendarToShare + ":\Calendar")

Start-Sleep -Seconds 4