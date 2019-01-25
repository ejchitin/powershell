# Script to Remove User.
# Version 3.0
# Evert J Hernandez
# 2/27/2018

####################################################### 
# Color Functions
function Receive-OutputG 
{
 process { Write-Host $_ -ForegroundColor Green }
} 

function Receive-OutputR 
{
 process { Write-Host $_ -ForegroundColor Red }
} 

function Receive-OutputY 
{
 process { Write-Host $_ -ForegroundColor Yellow }
} 

# Variables
$startDTM = (Get-Date)
$ADUser = Read-Host "Provide the AD User UPN logon"
$ADUserDisplayName = (Get-ADUser -Identity $ADUser -Properties displayName).displayName
$ADUserPassword = Read-Host "Provide new password" -AsSecureString
$DriveFullPath = "\\atl-ntap-cifs\IT\Helpdesk\UserRemove"
$FullPath = $DriveFullPath, $ADUserDisplayName -join "\"
$FTPPath = "\\KMS001\k2ftp"
$FTPFullPath = $FTPPath, $ADUser -join "\"
$FTPArchiveFolder = $FullPath, "FTP Files" -join "\"
$FDrive = "\\kf003"
$FADUser = $ADUser, "$" -join ""
$FFolder = $FDrive, $FADUser -join "\"
$FArchiveFolder = $FullPath, "F Drive Files" -join "\"
$FGroupsFolder = $FullPath, "AD Groups" -join "\"

# CSV and Log File Location
$path = "C:\ADScriptLogs\"
$UserSpecificFolder = $path, $ADUser -join "\"
$CSVPath = $UserSpecificFolder + "\Groups.csv"
$logfile = $UserSpecificFolder + "\logfile.txt"

# Creates Logs Folder
Write-Host "Looking for Logs Folder" -ForegroundColor Yellow
If(!(test-path $path))
{
New-Item -ItemType Directory -Force -Path $path
Write-Host "Folder not Found Creating..." -ForegroundColor Yellow
}
else
{
Write-Host "Folder Found" -ForegroundColor Green
}

#Creating User Logs Specific Folder
Write-Host "Creating user specific logs folder" -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path $UserSpecificFolder


# Creates log timestamp
(get-date).DateTime | add-content -path $logfile

# Import AD Module
Import-Module ActiveDirectory
Write-Output "Connected to AD" | add-content $logfile -passthru | Receive-OutputG

# Saving Member Of List to Log File
Write-Output "Exporting groups to CSV" | Add-Content $logfile -PassThru | Receive-OutputY
(Get-ADUser –Identity $ADUser –Properties MemberOf).MemberOf -replace '^CN=([^,]+),OU=.+$','$1' | Out-File $CSVPath
#Write-Output "MemerOf List" | add-content $logfile -passthru | Receive-OutputY
#Write-Output "$memberof" | add-content $logfile -passthru | Receive-OutputY

# Reset User Password
Write-Output "Reseting User Password" | add-content $logfile -passthru | Receive-OutputY
Set-ADAccountPassword -Identity $ADUser -NewPassword $ADUserPassword -Reset

# Retrieve the user in question:
$User = Get-ADUser $ADUser -Properties memberOf

# Retrieve groups that the user is a member of
$Groups = $User.memberOf |ForEach-Object {
    Get-ADGroup $_
} 

# Go through the groups and remove the user
Write-Output "Removed user from AD Groups" | add-content $logfile -passthru | Receive-OutputG
$Groups |ForEach-Object { Remove-ADGroupMember -Identity $_ -Members $User -Confirm:$false }

# Add User to ArchiveRestriction-USERS group
Write-Output "Added User to ArchiveRestriction-USERS group" | add-content $logfile -passthru | Receive-OutputG
Add-ADGroupMember -Identity "ArchiveRestriction-USERS" -Members $ADUser

# Move user to Deleted Users OU
Write-Output "Moving User to Deleted Users OU" | add-content $logfile -passthru | Receive-OutputY
Get-ADUser $ADUser | Move-ADObject -TargetPath 'OU=Deleted Users,DC=KDNS001,DC=LOCAL'

# Deny Access on the Dial-in Tab
Write-Output "Denying Access on the Dial-in Tab" | add-content $logfile -passthru | Receive-OutputY
Set-ADUser $ADUser -replace @{msNPAllowDialIn=$FALSE}

# Hide from addreess list
Write-Output "Hide User from addreess list" | add-content $logfile -passthru | Receive-OutputY
$user = Get-ADUser $ADUser –properties *
$user.msExchHideFromAddressLists = “True”
Set-ADUser –instance $user

# Connect to Office 365
Write-Output "Connecting to O365" | add-content $logfile -passthru | Receive-OutputG
$credential = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking


# Disable ActiveSync, OWA for devices and Outlook Web App
Write-Output "Disabled ActiveSync, OWA for devices and Outlook Web App" | add-content $logfile -passthru | Receive-OutputY
Set-CASMailbox -Identity $ADUser -ActiveSyncEnabled $False
Set-CASMailbox -Identity $ADUser -OWAforDevicesEnabled $false
Set-CASMailbox -Identity $ADUser -OWAEnabled $false

# Convert To Shared Mailbox
Write-Host "Converting user mailbox to Share-Mailbox..." -NoNewline | Add-Content $logfile | Receive-OutputG
Set-Mailbox -Identity $ADUser -Type shared
sleep -Seconds 5

#Connect MsolService to remove license
Write-Output "Removing O365 license" | Add-Content $logfile -PassThru | Receive-OutputY
Connect-MsolService -Credential $credential
$ADDomainUser = $ADUser + "@kidsii.com"
Set-MsolUserLicense -UserPrincipalName $ADDomainUser -RemoveLicenses "KidsII:ENTERPRISEPACK"

# Discconect from O365
Write-Output "Discconected from O365" | add-content $logfile -passthru | Receive-OutputG
$osession = Get-PSSession | select -expandProperty Name
Remove-PSSession -Name $osession

# Create Archive Folders
Write-Output "Created Archive Folders" | add-content $logfile -passthru | Receive-OutputG
New-Item $FullPath -type directory
New-Item $FTPArchiveFolder -type directory
New-Item $FArchiveFolder -type directory
New-Item $FGroupsFolder -Type Directory

# Copy files from FTP and F Drive to Archive Folder
Write-Output "Copying files from FTP and F Drive to Archive Folders" | add-content $logfile -passthru | Receive-OutputY
Copy-Item $FTPFullPath $FTPArchiveFolder -recurse
Copy-Item $FFolder $FArchiveFolder -recurse
Move-Item $CSVPath $FGroupsFolder -Force
Write-Output "Done copying files" | add-content $logfile -passthru | Receive-OutputG

# Remove F Drive and FTP Folders
Remove-Item -Path $FTPFullPath -Force -Recurse
Remove-Item -Path $FFolder -Force -Recurse

# Rename Log File
Write-Output "Renaming Log File to include Date and Time" | add-content $logfile -passthru | Receive-OutputY
Write-Output "End off Script" | add-content $logfile -passthru | Receive-OutputG
$TimeStamp = Get-Date | foreach {$_ -replace ":", "."} 
$TimeStampName = $TimeStamp | foreach {$_ -replace "/", "."} 
Rename-Item -Path $logfile -NewName "$ADUser Removal Log File $TimeStampName.txt"


# The End!