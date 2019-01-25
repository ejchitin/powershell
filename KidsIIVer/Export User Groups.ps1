Import-Module ActiveDirectory

$KidsIIUSer = Read-Host "Enter username"

#Cannot use the Export-Csv, because it will not give you the correct text. So Out-File is necessary

(Get-ADUser –Identity $KidsIIUSer –Properties MemberOf).MemberOf -replace '^CN=([^,]+),OU=.+$','$1' | Out-File "c:\Scripts\0101.csv"


