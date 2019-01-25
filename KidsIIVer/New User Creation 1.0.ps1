#  New User Creation

#  Evert J Hernandez

#  3/12/18

# Specials Thanks to: Matthew F and Eduardo N

#############################################################

# Connect to Local Exchange server
$AdminCredentials = Get-Credential -Message "Type your Admin Credentials to Connect to Local Exchange Server"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ATL-COR-EXCH-01.KDNS001.LOCAL/PowerShell/ -Authentication Kerberos -Credential $AdminCredentials
Import-PSSession $Session


# Variables & Functions
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null #Alouds the creation of a dialog box using VB
$ScriptDate = (Get-Date)
$OU = "OU=New Users,OU=KIDSII Users,DC=KDNS001,DC=LOCAL"
$Credentials = Get-Credential -Message "The username must be in this format: FirstName.LastName"
$UserName = $Credentials.UserName 
$UserPassword = $Credentials.Password
$UserFirstName = $UserName.Split(".")[0]
$UserLastName = $UserName.Split(".")[1]
$UserDisplayName = ($UserFirstName + " " + $UserLastName)
$UserCompany = "Kids II, Inc."
$UserJobTitle = [Microsoft.VisualBasic.Interaction]::InputBox("Enter user job title", "Job Title")
$UserDepartment = [Microsoft.VisualBasic.Interaction]::InputBox("Enter user department", "Department")
$Manager = [Microsoft.VisualBasic.Interaction]::InputBox("Enter user hiring manager Ex: Last Name, First Name", "Manager") #VB code for a dialog box that accepts inputs
$EmployeeID = [Microsoft.VisualBasic.Interaction]::InputBox("Enter employee ID", "Employee ID")

# Finding the correct manager. Thanks! to Matthew F.

if ($Manager -like "*.*"){  #Exception for name format like "Christi B. West"
                $ManagerLastname = $Manager.Split(", ")[0]
                $ManagerFirstname = $Manager.Split(",")[1]}
            else{
                $ManagerLastname = $Manager.Split(", ")[0]
                $ManagerFirstname = $Manager.Split("")[-1]}
            #Then create different possible combinations to use in the $ManagerDN Search
            $ManagerDN1 = "$ManagerFirstname" + " " + "$ManagerLastname"
            $ManagerDN2 = "$ManagerLastname" + " $ManagerFirstname"
            $ManagerDN3 = "$ManagerLastname" + "$ManagerFirstname"
            $ManagerDN4 = "*" + "$ManagerFirstName" + " " + "$ManagerLastname" #Exception for name format like "John Scott Randall"
            #Convert names to lower case. Appears to be case sensitive
            $ManagerLC = $Manager.ToLower()
            $ManagerDNLC1 = $ManagerDN1.ToLower()
            $ManagerDNLC2 = $ManagerDN2.ToLower()
            $ManagerDNLC3 = $ManagerDN3.ToLower()
            $ManagerDNLC4 = $ManagerDN4.ToLower()
            
            
                #MANAGER AD MAPPING --- find AD object for desired manager in import file
                Write-Host "Performing Manager lookup against AD..." -NoNewline -ForegroundColor Green -BackgroundColor Black  #Used different possible displaynames to search for a managername
                Try { $ManagerDN = IF ($Manager -ne ''){ 
                    (Get-ADUser -Credential $AdminCredentials -Filter {(Name -like $Manager) -or (Name -like $ManagerDN1) -or (
                    Name -like $ManagerDN2) -or (Name -like $ManagerDN3) -or (Name -like $ManagerLC) -or (Name -like $ManagerDNLC1) -or (
                    Name -like $ManagerDNLC2) -or (Name -like $ManagerDNLC3) -or (Name -like $ManagerDN4) -or (Name -like $ManagerDNLC4) -or ((GivenName -like $ManagerFirstname) -and (Surname -like $ManagerLastname))}).DistinguishedName} } #Manager required in DN format 
                Catch { }
                if ($ManagerDN){ 
                Write-Host "...FOUND: " $ManagerDN -ForegroundColor Green -BackgroundColor Black
                }




# Create New User
New-RemoteMailbox -Name $UserDisplayName `
                  -Password $UserPassword `
                  -FirstName $UserFirstName `
                  -LastName $UserLastName `
                  -DisplayName $UserDisplayName `
                  -UserPrincipalName ($UserName + "@kidsii.com") `
                  -OnPremisesOrganizationalUnit $OU

Write-Host "New User and Mailbox Created" -ForegroundColor Green -BackgroundColor Black 


# Remove PSSession
Write-Host "Terminating Exchange Session..." -NoNewline -ForegroundColor Green -BackgroundColor Black
$osession = Get-PSSession | select -expandProperty Name
Remove-PSSession -Name $osession
sleep -Seconds 5

#Connecting to Office 365 To Add remaining fields
<#Write-Host "Terminated. Connecting to Office 365" -ForegroundColor Green -BackgroundColor Black
$CloudCredentials = Get-Credential -Message "Enter your Office 365 credentials"
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $CloudCredentials -Authentication Basic -AllowRedirection
Import-PSSession $ExchangeSession -DisableNameChecking



# Connect MsolService to add license
# Cannot add licenses from this script because user will not be found on the cloud so quick after creation.


# Removing O365 Session
Write-Host "Terminating O365 session" -ForegroundColor Green -BackgroundColor Black
$o365session = Get-PSSession | select -expandProperty Name
Remove-PSSession -Name $o365session #>

#Fill remaining user data
Import-Module ActiveDirectory
Write-Host "Finalizing user cloud setup" -ForegroundColor Green -BackgroundColor Black
#Set-ADUser -Identity $UserName -Replace @{company="$UserCompany";Title=$UserJobTitle;Department=$UserDepartment;Manager=$ManagerDN;EmployeeID=$EmployeeID} 
Get-ADUser -Identity "$UserName" | Set-ADUser -Company "$UserCompany" -Title "$UserJobTitle" -Department "$UserDepartment" -Manager "$ManagerDN" -EmployeeID "$EmployeeID"

# Emulating User
                 ###Copy AD Groups
  
    $Source = [Microsoft.VisualBasic.Interaction]::InputBox("Enter username to emulate", "Emulation")
    $Target = $UserName
Write-Host "Copying AD Groups" -ForegroundColor Green -BackgroundColor Black

# Retrieve group memberships.
    $SourceUser = Get-ADUser $Source -Properties memberOf
    $TargetUser = Get-ADUser $Target -Properties memberOf

# Hash table of source user groups.
    $List = @{}

#Enumerate direct group memberships of source user.
    ForEach ($SourceDN In $SourceUser.memberOf)
        {
    # Add this group to hash table.
            $List.Add($SourceDN, $True)
    # Bind to group object.
            $SourceGroup = [ADSI]"LDAP://$SourceDN"
    # Check if target user is already a member of this group.
            If ($SourceGroup.IsMember("LDAP://" + $TargetUser.distinguishedName) -eq $False)
                {
        # Add the target user to this group.
                    Add-ADGroupMember -Identity $SourceDN -Members $Target
                    
                }
        }

# Enumerate direct group memberships of target user.
    ForEach ($TargetDN In $TargetUser.memberOf)
        {
    # Check if source user is a member of this group.
            If ($List.ContainsKey($TargetDN) -eq $False)
                {
        # Source user not a member of this group.
        # Remove target user from this group.
                    Remove-ADGroupMember $TargetDN $Target
                }
        }

Write-Host "Emulation Completed" -ForegroundColor Green -BackgroundColor Black


# End Script

Write-Host "End of script (EoS), Please check results" -ForegroundColor Green -BackgroundColor Black








