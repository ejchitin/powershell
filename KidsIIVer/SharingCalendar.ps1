Write-Host "Getting O365 Credentials" -ForegroundColor Green -BackgroundColor Black
Start-Sleep -Seconds 2


$credentials = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Write-Host "Enter email of calendar to share" -ForegroundColor Green -BackgroundColor Black
$calendarToShare = Read-Host

Write-Host "Enter user that needs access" -ForegroundColor Green -BackgroundColor Black
$oUser = Read-Host

Write-Host "Sharing calendar" -ForegroundColor Green -BackgroundColor Black
Add-MailboxFolderPermission -Identity ($calendarToShare + ":\Calendar") -AccessRights Owner -User $oUser 

Start-Sleep -Seconds 4