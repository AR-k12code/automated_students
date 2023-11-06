#Requires -Version 7.0
#Get-Help .\automated_students.ps1
#Get-Help .\automated_students.ps1 -Examples
#  ___ _____ ___  ___   ___   ___    _  _  ___ _____         
# / __|_   _/ _ \| _ \ |   \ / _ \  | \| |/ _ \_   _|        
# \__ \ | || (_) |  _/ | |) | (_) | | .` | (_) || |          
# |___/ |_| \___/|_|   |___/ \___/  |_|\_|\___/ |_|

#  ___ ___ ___ _____   _____ _  _ ___ ___   ___ ___ _    ___ 
# | __|   \_ _|_   _| |_   _| || |_ _/ __| | __|_ _| |  | __|
# | _|| |) | |  | |     | | | __ || |\__ \ | _| | || |__| _| 
# |___|___/___| |_|     |_| |_||_|___|___/ |_| |___|____|___|
#    
# Please see the https://www.github.com/carbm1/automated_students for more information.

<#
  .SYNOPSIS
  This script is used to manage student accounts in Active Directory based on the database built by automated_database.ps1.
  
  .NOTES
    Author: Craig Millsap
    Creation Date: 3/2021
    Note: We should probably turn on Indexing on the EmployeeNumber attribute in AD.

  .EXAMPLE
  PS> .\automated_students.ps1 -Staging
  This will write to terminal all the changes the script believes it will make to Active Directory. Including creating OU's, new students, groups, etc.

  .EXAMPLE
  PS> .\automated_students.ps1 -QuickMode
  This will compare SQLite and Active Directory to find mismatched Student ID, First Name, Last Name. This will skip all other students.
  
  .PARAMETER Staging
  This parameter is used to control if you want changes made to your AD or not. It is ON by default in the settings.ps1 file.

  .PARAMETER VerboseStudent
  This will output to the terminal information about existing and new students while the script is running.

  .PARAMETER QuickMode
  This mode will compare the students in the SQLite Database and Active Directory and only evaluate the mismatched Student ID, First Name, and Last Name.

  .PARAMETER SkipStudents
  Skips processing existing or new students. This will still create the OU structure and groups.

  .PARAMETER SkipNewStudents
  Skip creating new student acounts. This will still process existing accounts.

  .PARAMETER SkipExistingStudents
  Skip processing existing student accounts. This will still create new accounts.

  .PARAMETER StopAfterXNew
  This will stop the script after it creates X number of new students.

  .PARAMETER StopAfterXExisting
  This will stop the script after it processes X number of existing students.

  .PARAMETER DisablePostProcessingScript
  This will disable invoking x_PostProcessingAutomatedStudents.ps1. Staging also disables this.

#>

Param(
    [Parameter(Mandatory=$false)][switch]$Staging, #don't make changes to domain
    [Parameter(Mandatory=$false)][switch]$SkipStudents,
    [Parameter(Mandatory=$false)][switch]$SkipNewStudents,
    [Parameter(Mandatory=$false)][switch]$SkipExistingStudents,
    [Parameter(Mandatory=$false)][switch]$VerboseStudent,
    [Parameter(Mandatory=$false)][switch]$DisablePostProcessingScript, #Do not run the final script.
    [Parameter(Mandatory=$false)][switch]$RenewEmailPassword,
    [Parameter(Mandatory=$false)][int]$StopAfterXNew, #Testing something on new accounts? Only create this number then quit. Useful while testing x_InterimProcessingNewAccounts.ps1
    [Parameter(Mandatory=$false)][int]$StopAfterXExisting, #Testing on existing accounts? Only create this number then quit. Useful while testing x_InterimProcessingExistingAccounts.ps1
    [Parameter(Mandatory=$false)][switch]$ForceDoNotRequireDomainAdmin, #I'm serious. You can mess things up if you can't read confidentiality bits.
    [Parameter(Mandatory=$false)][switch]$SkipDisablingAccounts, #This will skip disabling accounts that have an EmployeeNumber and are not in the exclusions or in the students table. This does nothing for accounts that do not have EmployeeNumbers.
    [Parameter(Mandatory=$false)][switch]$QuickMode #This will only evaulate accounts that do not have a match on a Student ID in AD AND have a mismatched GivenName and Surname in case they neeed reanmed. All other accounts are ignored.
)

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
if (-Not(Test-Path $currentPath\logs)) { New-Item -ItemType Directory -Path $currentPath\logs }
$logfile = "$currentPath\logs\automated_students_$(get-date -f yyyy-MM-dd-HH-mm-ss).log"
try {
    Start-Transcript($logfile)
} catch {
    Stop-TranScript; Start-Transcript($logfile)
}

$timestamp = [int64](Get-Date -UFormat %s)

#Pull in settings
if (Test-Path $currentPath\settings.ps1) {
    . $currentPath\settings.ps1
} else {
    Write-Host "Error: Missing settings.ps1 file. Please read documentation." -ForegroundColor Red
    exit(1)
}

#Pull in functions
if (Test-Path $currentPath\z_functions.ps1) {
    . $currentPath\z_functions.ps1
} else {
    Write-Host "Error: Missing z_functions.ps1 file. Please read documentation." -ForegroundColor Red
    exit(1)
}

if ($Staging) {
    Write-Host "Info: Staging flag has been specified. No changes will be made to the domain." -ForegroundColor Yellow
    Write-Host "Info: Staging will NOT run x_PostProcessingAutomatedStudents.ps1. It will run x_InterimProcessingExistingAccounts.ps1 and x_InterimProcessingNewAccounts.ps1. You must account for `$staging in your code!" -ForegroundColor Yellow
}

#####################################################################
# Must be ran as an administrator. This is primarily because of
# NTFS permissions but other things fail as well sometimes.
#####################################################################

if (-Not($DoNotRequireDomainAdmin)) { #Make sure you know what you're doing with confidentiality bits and creating a role that can read them.
    if (-Not $(New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole("Domain Admins")) {
        Write-Host "Error: This script must be run by a Domain Admin account."
        exit(1)
    }
} else {
    Write-Host "Info: You've specified to not require domain admin rights. Be sure the user account can read confidential fields in AD."
    if (-Not($ForceDoNotRequireDomainAdmin)) {
        if ((Get-ADUser -Filter { EmployeeNumber -like "*" } -Properties EmployeeNumber | Measure-Object).Count -eq 0) {
            Write-Host "Error: 0 accounts returned from AD with the confidential EmployeeNumber field. Unless you specify -ForceDoNotRequireDomainAdmin this script will not run to save you from yourself."
            exit(1)
        }
    }
}

if (-Not($DoNotRequireAdminstrator)) { #Hope you know what you're doing.
    if (-Not $(New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Error: Must run as administrator!"
        exit(1)
    }
}

#####################################################################
# Dependencies - ActiveDirectory, NTFSSecurity,PSQLite Modules
#####################################################################

if (Get-Module -ListAvailable | Where-Object {$PSItem.name -eq "ActiveDirectory"}) {
    try { Import-Module ActiveDirectory } catch { Write-Host "Error: Unable to load module ActiveDirectory." -ForegroundColor RED; exit(1) }
} else {
    Write-Host 'ActiveDirectory Module not found!'

    if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically install it? y/n')) {
        try {
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
            Restart-Service wuauserv
            DISM.exe /Online /add-capability /CapabilityName:Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
            Restart-Service wuauserv
            Import-Module ActiveDirectory
        } catch {
            Write-Host "Failed to install ActiveDirectory Module. Please see https://docs.microsoft.com/en-us/windows-server/remote/remote-server-administration-tools"
            exit(1)
        }
    } else {
        exit(1)
    }
}

if (Get-Module -ListAvailable | Where-Object {$PSItem.name -eq "NTFSSecurity"}) {
    try { Import-Module NTFSSecurity } catch { Write-Host "Error: Unable to load module NTFSSecurity." -ForegroundColor RED; exit(1) }
} else {
    Write-Host 'NTFSSecurity Module not found!'
    if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically install it? y/n')) {
        try {
            Install-Module -Name NTFSSecurity -Scope AllUsers -Force
            Import-Module NTFSSecurity
        } catch {
            Write-Host 'Failed to install NTFSSecurity Module.'
            exit(1)
        }
    } else {
        exit(1)
    }
}

#SimplySQL Module is required.
if (Get-Module -ListAvailable | Where-Object {$PSItem.name -eq "SimplySQL"}) {
    Import-Module SimplySQL
  } else {
    Write-Host 'SimplySQL Module not found!'
    if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically install it? y/n')) {
        try {
            Install-Module -Name SimplySQL -Scope AllUsers -Force
            Import-Module SimplySQL
        } catch {
            Write-Host 'Failed to install SimplySQL Module.' -ForegroundColor RED
            exit(1)
        }
    } else {
        exit(1)
    }
  }

#Connect Database
try {
    Connect-Database -Database $database
} catch {
    Write-Host "Error: Failed to connect to database engine." -ForegroundColor Red; exit(1)
}

#mail
if ($smtpAuth) {
    if ((Test-Path ($smtpPasswordFile)) -and -Not($RenewEmailPassword)) {
        $smtpPassword = Get-Content $smtpPasswordFile | ConvertTo-SecureString
    } elseif ($RenewEmailPassword) {
        Write-Host "Info: SMTP Auth password renewal." -ForeGroundColor Yellow
        Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $smtpPasswordFile
        $smtpPassword = Get-Content $smtpPasswordFile | ConvertTo-SecureString
    } else {
        Write-Host "Info: SMTP Password file does not exist! [$smtpPasswordFile]. Please enter the password to be saved on this computer for emails." -ForeGroundColor Yellow
        Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $smtpPasswordFile
        $smtpPassword = Get-Content $smtpPasswordFile | ConvertTo-SecureString
    }
    $mailCredentials = New-Object -Type System.Management.Automation.PSCredential -ArgumentList $sendMailFrom, $smtpPassword
}

#due to samaccountname limitations we can't have a shortname of a school be longer than 9 characters.
$validschools.Values | ForEach-Object {
    if ($PSItem.Length -eq 0 -or $PSItem.Length -gt 9) {
        Write-Host 'Error: Valid Schools ShortNames exceed maximum 9 characters.'
        exit(1)
    }
}

#Active Directory Information
try {
    $domain = $(Get-ADDomain).DistinguishedName
    $domainfqdn = $(Get-ADDomain).Forest
    
    #This is necessary in a large domain. You need to work with the operations managers of your domain controllers or cmdlets timeout.
    $PSDefaultParameterValues = @{"*-AD*:Server"="$((Get-ADDomain).InfrastructureMaster)"}

} catch {
    Write-Host 'Error: Unable to Access domain controller.' -ForegroundColor Red
    Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Unable to Access domain controller."
    exit(1)
}

#Verify AD Replication
$ADServersOutOfSync = @()
Get-ADReplicationPartnerMetadata -Target "$env:userdnsdomain" -Scope Domain | Select-Object Server, LastReplicationAttempt, LastReplicationSuccess | ForEach-Object {
    if ($PSItem.LastReplicationAttempt -ne $PSItem.LastReplicationSuccess) { 
        Write-Host "WARN: There appears to be an AD Replication Issue with",$PSItem
        $ADServersOutOfSync += $PSItem
    }
}
if ($ADServersOutOfSync.Count -ge 1) {
    Send-EmailNotification -subject "Automated_Students: AD Replication Error" -body "Error: There are AD replications that have not succeeded The Automated Students will contintue but these replication failures can cause other issues.`r`nPlease review:`r`n`
    $($ADServersOutOfSync | ForEach-Object { "Server: $($PSitem.Server)`r`nLast Replication Attempt: $($PSitem.LastReplicationAttempt)`r`nLast Replication Success: $($PSitem.LastReplicationSuccess)`r`n`r`n" })"
}

#Find GAM if requested.
if (($GAMprecreateUser) -OR ($GAMsetModerationRules) -OR ($GAMVerifyModerationRules)) {
    Write-Host "Info: GAM Moderation and user creation specified in settings.ps1. Verifying configuration and domain access."
    if (Test-Path $currentPath\..\gam\gam.exe) {
        & $currentPath\..\gam\gam.exe info domain
        if ($LASTEXITCODE -ge 1) {
            Write-Host "Error: You have specified that you want to use GAM to manage G-Suite. However it appears to not be configured correctly." -ForegroundColor RED
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Please see the readme about configuring GAM correctly."
            exit(1)
        }
    } else {
        Write-Host "Error: You have specified that you want to use GAM to manage G-Suite. However it is not where this script expects it to be at c:\scripts\gam\gam.exe" -ForegroundColor RED
        Send-EmailNotification -subject "Automated_Students: Failure" -body "Please see the readme about configuring GAM correctly."
        exit(1)
    }
}

#####################################################################
# Pull Students data from the SQLite database.
#####################################################################
try {
    #$studentsCSV = invoke-SqlQuery -Query "SELECT students.*,Student_nickname,Student_gradyr FROM students LEFT JOIN students_extras ON students.Student_id = students_extras.Student_id ORDER BY Student_id DESC" -ErrorAction 'STOP' 
    #sorting by highest number first causes new accounts to be created first BUT it if you redo things it will cause a younger duplicate name conflict to have precidence.
    $studentsCSV = invoke-SqlQuery -Query 'SELECT
        students.*,Student_nickname,Student_gradyr
        FROM students
        LEFT JOIN students_extras ON students.Student_id = students_extras.Student_id
        ORDER BY Student_id' -ErrorAction 'STOP'
    $validStudentIDs = $studentsCSV | Select-Object -ExpandProperty 'Student_id'
    if (($validStudentIDs | Measure-Object).Count -lt $minStuCount) {
        throw 'Error: Not enough students returned from database. Check $minStuCount variable in settings.ps1.'
    } else {
        Write-Host "Info: Found" (($validStudentIDs | Measure-Object).Count) "students in the database to process."
    }
} catch {
    Write-Host "Error: Failed to query minumum student count from database. Exiting."
    Send-EmailNotification -subject "Automated_Students: Failure" -body "Failed to query minumum student count from database. Exiting."
    exit(1)
}

# Verify we do not have any duplicate EmployeeNumbers in AD. This WILL break things if there are duplicate EmployeeNumbers anywhere in the directory.
try {
    $duplicateCheck = Get-ADUser -Filter { EmployeeNumber -like "*" } -Properties EmployeeNumber | Group-Object -Property EmployeeNumber | Where-Object { $PSItem.count -gt 1 }
    if ($duplicateCheck.Count -gt 1) {
        $duplicateCheck | ForEach-Object {
            $PSItem.Group | ForEach-Object {
                Write-Host "Error: $($PSItem.EmployeeNumber),$($PSItem.DistinguishedName)"
            }
        }
        Write-Host "Error: There are duplicate EmployeeNumber attributes in your Active Directory. You must resolve this before continuing." -ForeGroundColor RED
        Send-EmailNotification -subject "Automated_Students: Duplicate Failure" -body "Error: There are duplicate EmployeeNumber attributes in your Active Directory. You must resolve this before continuing.`r`nPlease inspect the log file $($logfile) for more details."
        exit(1)
    }
} catch { 
    Write-Host "Error: Could not run duplicate EmployeeNumber check on domain."
    exit (1)
}

#####################################################################
# Find students who are Enabled in AD but are not in the CSV.
# Loop throught the OUs based on $adStructure, compare, then
# disable the accounts.
#####################################################################

$stuOUs = @()
switch ($adStructure) {
    1 { $stuOUs = @("ou=Students,$($domain)") }
    2 { $stuOUs = $validschools.Values | ForEach-Object { $stuOUs += @("ou=$($PSItem),ou=Students,$($domain)") } }
    3 { $stuOUs = $validschools.Values | ForEach-Object { $stuOUs += @("ou=Students,ou=$($PSItem),$($domain)") } }
    4 { $stuOUs = @("ou=Students,$($domain)") }
    5 { $stuOUs = $validschools.Values | ForEach-Object { $stuOUs += @("ou=$($PSItem),ou=Students,$($domain)") } }
    6 { $stuOUs = $validschools.Values | ForEach-Object { $stuOUs += @("ou=Students,ou=$($PSItem),$($domain)") } }
}

#get currently active student ID numbers in AD
if (@(1,2,4,5) -contains $adStructure) {
    $adStudents = Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=Students,$domain" -Properties EmployeeNumber,memberof | Where-Object { $PSItem.DistinguishedName -notlike "*OU=Excluded*" }
} elseif (@(3,6) -contains $adStructure) {
    $adStudents = @()
    $stuOUs | ForEach-Object {
        $adStudents += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "$PSItem" -Properties EmployeeNumber,memberof | Where-Object { $PSItem.DistinguishedName -notlike "*OU=Excluded*" }
    }
}

$activeStudentIDs = @()
$activeStudentIDs += $adStudents | Select-Object -ExpandProperty EmployeeNumber

#Calculate changes and if it exceeds our daily maximum then abort.
if (-Not($SkipDisablingAccounts)) {
    #Count New and Deactivated.
    $calculatedChanges = (Compare-Object $activeStudentIDs $validStudentIDs -PassThru).count
    if ($calculatedChanges -ge $maxChanges) {
        Write-Host "Error: Calculated changes",$calculatedChanges,"exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. Disabling accounts are included in this count."
        if (-Not($staging)) {
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Calculated changes $($calculatedChanges) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. Disabling accounts are included in this count."
            exit(1)
        }
    } else {
        Write-Host "Info: Current allowed number of new and disabled students is $($maxChanges). This is roughly $([int]($maxChanges / ($studentsCSV | measure-object).Count * 100))% of your students."
        Write-Host "Info: Calculated number of new/disabled accounts:",$calculatedChanges
    }
} else {
    #Count New Only
    $calculatedNewChanges = (Compare-Object $activeStudentIDs $validStudentIDs | Where-Object { $PSItem.SideIndicator -eq '=>' }).count
    if ($calculatedNewChanges -ge $maxChanges) {
        Write-Host "Error: Calculated changes $($calculatedNewChanges) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. The disabling of accounts are NOT included in this count."
        if (-Not($staging)) {
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Calculated changes $($calculatedNewChanges) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. The disabling of accounts are NOT included in this count."
            exit(1)
        }
    } else {
        Write-Host "Info: Current allowed number of new and disabled students is $($maxChanges). This is roughly $([int]($maxChanges / ($studentsCSV | measure-object).Count * 100))% of your students."
        Write-Host "Info: Calculated number of new accounts:",$calculatedNewChanges
    }
}

#Exclusion List
if (Test-Path $currentPath\exclusions.csv) {
    Write-Host "Info: Exclusion CSV detected." -ForegroundColor YELLOW
    $excludeCSV = Import-CSV $currentPath\exclusions.csv
    #$headers = 
    if ((Get-Content $currentPath\exclusions.csv).Count -gt 1) {
        if (($excludeCSV | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -contains 'Student_id') {
            $excludedStudentIDs = $excludeCSV | Select-Object -ExpandProperty Student_id
            #Now we need their email addresses to exclude them from Distribution Group Removal.
            $excludedAccounts = @()
            $excludedStudentIDs | ForEach-Object { $id = $PSItem; $excludedAccounts += Get-ADUser -Filter { EmployeeNumber -eq $id } | Select-Object -ExpandProperty SamAccountName }
            Write-Host "Info: Student ID numbers", "$($excludedStudentIDs -join (','))", "will be excluded from this automation script. Account will not be created, modified, or disabled." -ForegroundColor YELLOW
        } else {
            Write-Host "Error: Exclusion CSV does not contain a Student_id column!" -ForegroundColor RED
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Exclusion CSV does not contain a Student_id column."
            exit(1)
        }
    }
}

if (-Not($SkipDisablingAccounts)) {
    Compare-Object $activeStudentIDs $validStudentIDs -PassThru | Where-Object { $PSItem.'SideIndicator' -eq '<=' } | ForEach-Object {
        $studentid = $PSItem

        #exclude closing account
        if ($excludedStudentIDs -contains $studentid) { return }

        #select individual student without another ad query
        $accountToDisable = $adStudents | Where-Object { $PSItem.EmployeeNumber -eq $studentid }
        #disable the account
        if ($staging) { 
            Write-Host "Staging: Disabling account $($accountToDisable.'DistinguishedName')" -ForeGroundColor Yellow
            Disable-ADAccount -Identity $accountToDisable -WhatIf
        } else {
            Write-Host "Info: Disabling account $($accountToDisable.'DistinguishedName')" -ForeGroundColor Yellow
            Disable-ADAccount -Identity $accountToDisable -Confirm:$False
            
            #Log Event Disable
            Invoke-SqlUpdate -Query "INSERT INTO action_log (Student_id, Identity, Action, Timestamp) VALUES ($studentid,""$($accountToDisable.UserPrincipalName)"",""Disable"",""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"")" | Out-Null

            #remove from all group memberships
            $accountToDisable | Select-Object -ExpandProperty memberof | ForEach-Object { Remove-ADPrincipalGroupMembership -Identity $accountToDisable -MemberOf $PSItem -Confirm:$false }
        }
    }
} else {
    Write-Host "Info: You have turned off disabling existing accounts with the `$SkipDisablingAccounts in settings.ps1." -ForegroundColor Yellow
    if ($staging) {
        Write-Host "Info: Staging will not report accounts that would be disabled with `$SkipDisablingAccounts enabled." -ForegroundColor Yellow
    }
}

#Overrides List
if (Test-Path $currentPath\overrides.csv) {
    Write-Host "Info: Overrides CSV detected." -ForegroundColor YELLOW
    $overridesCSV = Import-CSV $currentPath\overrides.csv
    #$headers = 
    if ((Get-Content $currentPath\overrides.csv).Count -gt 1) {
        if (($overridesCSV | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -contains 'Student_id') {
            #Student_id needs to be an integer instead of string.
            $overridedAccounts = $overridesCSV | Select-Object -Property @{Name = 'Student_id'; Expression = { [int]$PSitem.'Student_id' }},First_name,Last_name,Middle_Initial | Group-Object -Property Student_id -AsHashTable
            Write-Host "Info: Student ID numbers", "$($overridedAccounts.Keys -join (','))", "will have their information overridden per the overrides.csv file." -ForegroundColor YELLOW
        } else {
            Write-Host "Error: Overrides CSV does not contain a Student_id column!" -ForegroundColor RED
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Override CSV does not contain a Student_id column."
            exit(1)
        }
    }
}

#QuickMode
#This will compare our AD to the SQL Database and only run against accounts that do not exist in AD or have mismatched GivenName and Surname.
if ($QuickMode) {
    try {
        $adstudents = Get-ADUser -Filter { Enabled -eq $True -and EmployeeNumber -like "*" } -SearchBase "ou=Students,$((Get-ADDomain).DistinguishedName)" -properties EmployeeNumber #| Select-Object -Property @{Name='EmployeeNumber';Expression={[int]$PSItem.'EmployeeNumber'}},GivenName,Surname
        $adstudentsobj = @()
        $adstudents | ForEach-Object { 
            #evaluate if overriddedAccounts has this student.
            if ($overridedAccounts.([int]($PSItem.EmployeeNumber))) {
                $adstudentsobj += [PSCUSTOMOBJECT]@{ "EmployeeNumber" = $PSItem.EmployeeNumber; "Surname" = ($overridedAccounts.([int]($PSItem.EmployeeNumber))).Last_name; "GivenName" = ($overridedAccounts.([int]($PSItem.EmployeeNumber))).First_name }
            } else {
                $adstudentsobj += [PSCUSTOMOBJECT]@{ "EmployeeNumber" = $PSItem.EmployeeNumber; "Surname" = $PSItem.surname; "GivenName" = $PSItem.givenName }
            }
        }

        #If $useNickname is set to true then evaluate on the nickname field for QuickMode.
        if ($useNickname) {
            $sqlstudentsobj = invoke-SqlQuery -Query 'SELECT
            students.Student_id AS EmployeeNumber,
            CASE
                WHEN LENGTH(Student_nickname) > 1 THEN Student_nickname
                ELSE First_name
                END
            AS Givenname,
            Last_name as Surname
            FROM students
            INNER JOIN students_extras ON students.Student_id = students_extras.Student_id'
        } else {
            $sqlstudentsobj = invoke-SqlQuery -Query "SELECT Student_id AS EmployeeNumber,First_name as GivenName,Last_name as Surname FROM students"
        }

        $studentIds = Compare-Object -ReferenceObject $adstudentsobj -DifferenceObject $sqlstudentsobj -Property EmployeeNumber,Surname,Givenname -PassThru | Select-Object -ExpandProperty EmployeeNumber -Unique
        
        if ($studentIds.Count -ge 1) {
            $studentsCSV = invoke-SqlQuery -Query "SELECT students.*,Student_nickname,Student_gradyr `
            FROM students `
            LEFT JOIN students_extras ON students.Student_id = students_extras.Student_id `
            WHERE students.Student_id IN ($($studentIds -join ',')) `
            ORDER BY students.Student_id" -ErrorAction 'STOP'
        } else {
            $studentsCSV = @()
        }

        Write-Host "Info: QuickMode has been specified. We will only be evaulating", ($studentIds | Measure-Object).count, "accounts."
    } catch {
        Write-Host "Error: Failed to compare and query students in QuickMode. $($_)"
        exit(1)
    }
}

#####################################################################
# Build the required organizational units in Active Directory
# depending on their $adStructure specification in settings.ps1
# Create management accounts for student password resets and 
# set the permissions on the OU
#####################################################################

#we have to create the stuOUs again with their Grad Year.
$stuOUs = @()
switch ($adStructure) {
    1 { #1 = STUDENTS/GRADYR
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Student_gradyr -Unique | Where-Object { $PSItem.Student_gradyr -ne $null }
        $stuOUs += @("ou=Students,$($domain)","ou=Disabled,ou=Students,$($domain)","ou=Restricted,ou=Students,$($domain)")
        $requiredOUs | ForEach-Object {
            $stuOUs += @("ou=$($PSItem.'Student_gradyr'),ou=Students,$($domain)")
        }
    }
    2 { #2 = STUDENTS/SCHOOL/GRADYR
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Student_gradyr -Unique | Where-Object { $PSItem.Student_gradyr -ne $null }
        $stuOUs += @("ou=Students,$($domain)","ou=Disabled,ou=Students,$($domain)")
        $validschools.Values | ForEach-Object {
            $stuOUs += @("ou=$($PSItem),ou=Students,$($domain)","ou=Restricted,ou=$($PSItem),ou=Students,$($domain)")
        }
        $requiredOUs | ForEach-Object {
            $stuOUs += @("ou=$($PSItem.'Student_gradyr'),ou=$($validschools.$([int]$PSItem.'School_id')),ou=Students,$($domain)")
        }
    }
    3 { #3 = SCHOOL/STUDENTS/GRADYR
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Student_gradyr -Unique | Where-Object { $PSItem.Student_gradyr -ne $null }
        $validschools.Values | ForEach-Object { $stuOUs += @("ou=$($PSItem),$($domain)") }
        $validschools.Values | ForEach-Object {
            $stuOUs += @("ou=Students,ou=$($PSItem),$($domain)","ou=Disabled,ou=Students,ou=$($PSItem),$($domain)","ou=Restricted,ou=Students,ou=$($PSItem),$($domain)")
        }
        $requiredOUs | ForEach-Object {
            $stuOUs += @("ou=$($PSItem.'Student_gradyr'),ou=Students,ou=$($validschools.$([int]$PSItem.'School_id')),$($domain)")
        }
    }
    4 { #4 = STUDENTS/GRADE
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Grade -Unique
        $stuOUs += @("ou=Students,$($domain)","ou=Disabled,ou=Students,$($domain)","ou=Restricted,ou=Students,$($domain)")
        $requiredOUs | ForEach-Object {
            
            $grade = $PSItem.'Grade'

            switch ($grade) {
                'Prekindergarten' { $grade = 'PK' }
                'Kindergarten' { $grade = 'K' }
            }

            $stuOUs += @("ou=$($grade),ou=Students,$($domain)")
        }
    }
    5 { #5 = STUDENTS/SCHOOL/GRADE
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Grade -Unique
        $stuOUs += @("ou=Students,$($domain)","ou=Disabled,ou=Students,$($domain)","ou=Restricted,ou=Students,$($domain)")
        $validschools.Values | ForEach-Object {
            $stuOUs += @("ou=$($PSItem),ou=Students,$($domain)")
        }
        $requiredOUs | ForEach-Object {
            
            $grade = $PSItem.'Grade'

            switch ($grade) {
                'Prekindergarten' { $grade = 'PK' }
                'Kindergarten' { $grade = 'K' }
            }

            $stuOUs += @("ou=$($grade),ou=$($validschools.$([int]$PSItem.'School_id')),ou=Students,$($domain)")
        }
    }
    6 { #6 = SCHOOL/STUDENTS/GRADE
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Grade -Unique
        $validschools.Values | ForEach-Object { $stuOUs += @("ou=$($PSItem),$($domain)") }
        $validschools.Values | ForEach-Object {
            $stuOUs += @("ou=Students,ou=$($PSItem),$($domain)","ou=Disabled,ou=Students,ou=$($PSItem),$($domain)","ou=Restricted,ou=Students,ou=$($PSItem),$($domain)")
        }
        $requiredOUs | ForEach-Object {
            
            $grade = $PSItem.'Grade'

            switch ($grade) {
                'Prekindergarten' { $grade = 'PK' }
                'Kindergarten' { $grade = 'K' }
            }

            $stuOUs += @("ou=$($grade),ou=Students,ou=$($validschools.$([int]$PSItem.'School_id')),$($domain)")
        }
    }
}

$currentOUs = Get-ADOrganizationalUnit -Filter * | Group-Object -Property DistinguishedName
$stuOUs | ForEach-Object {
    $ouString = $PSItem
    #Write-Host $ouString
    $ouName = $ouString.split(',')[0].split('=')[1]
    $ouPath = $ouString.split(',')[1..($ouString.Length-1)] -join(',')
    if (-Not($currentOUs | Where-Object { $PSItem.Name -eq $ouString })) {
        if ($staging) {
            Write-Host "Staging: New Organizational Unit at $($ouString)" -ForeGroundColor Yellow
            New-ADOrganizationalUnit -Name $ouName -Path $ouPath -ProtectedFromAccidentalDeletion $false -WhatIf
        } else {
            Write-Host "Creating Organizational Unit $($ouString)"
            try {
                New-ADOrganizationalUnit -Name $ouName -Path $ouPath -ProtectedFromAccidentalDeletion $false
            } catch {
                Write-Host $Error[0].Exception.GetType().fullname
                Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Could not create OU $ouString. Exiting."
                exit(1)                
            }
        }
    }
}

#Need to set permissions on the OU and assign a managment group for each school.
switch ($adStructure) {
    1 { 
        #permissions must be set on each graduation year
        $ou = "ou=Students,$($domain)"
        $requiredOUs | ForEach-Object {
            $groupName = "Student Management Accounts - $($validschools.$([int]$PSItem.'School_id'))"
            if (-Not(Get-AdGroup -Filter { SamAccountName -eq $groupName } -ErrorAction SilentlyContinue)) {
                if ($staging) {
                    Write-Host "Staging: New Security Group `"$($groupName)`" at $($ou)" -ForeGroundColor Yellow
                    New-ADGroup -Name $groupName -Path $ou -GroupScope Global -WhatIf
                } else {
                    New-ADGroup -Name $groupName -Path $ou -GroupScope Global
                }
            }
            $gradYearOU = "ou=$($PSItem.'Student_gradyr'),ou=Students,$($domain)"
            try {
                if ($staging) {
                    Write-Host "Staging: Grant Password Reset for `"$($groupName)`" on $($ou)" -ForeGroundColor Yellow
                } else {
                    Write-Host "Notify: Setting permissions for `"$groupName`" on $($gradYearOU)"
                    Grant-PasswordResetOnOU -group $groupName -ou $gradYearOU
                }
            } catch {
                Write-Host "Error: Failed to set permissions for the Management Account `"$groupName`" on $($gradYearOU)."
                Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Failed to set permissions for the Management Account `"$groupName`" on $($gradYearOU)."
                exit(1)
            }
        }
    }
    #2 and 5 are identical.
    { @(2,5) -contains $_ } {
        #permissions need to be set on the school OU only.
        $validschools.Values | ForEach-Object {
            $ou = "ou=$($PSItem),ou=Students,$($domain)"
            $groupName = "Student Management Accounts - $($PSItem)"
            if (-Not(Get-AdGroup -Filter { SamAccountName -eq $groupName } -ErrorAction SilentlyContinue)) {
                if ($staging) {
                    Write-Host "Staging: New Security Group `"$($groupName)`" at $($ou)" -ForeGroundColor Yellow
                    New-ADGroup -Name $groupName -Path "ou=Students,$($domain)" -GroupScope Global -WhatIf
                } else {
                    New-ADGroup -Name $groupName -Path "ou=Students,$($domain)" -GroupScope Global
                }
            }
            try {
                Write-Host "Setting permissions for `"$groupName`" on $($ou)"
                if ($staging) {
                    Write-Host "Staging: Grant Password Reset for `"$($groupName)`" on $($ou)" -ForeGroundColor Yellow
                } else {
                    Grant-PasswordResetOnOU -group $groupName -ou $ou
                }
            } catch {
                Write-Host "Error: Failed to set permissions for the Management Account `"$groupName`" on $($ou)."
                Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Failed to set permissions for the Management Account `"$groupName`" on $($ou)."
                exit(1)
            }
        }
    }
    #3 and 6 are identical
    { @(3,6) -contains $_ } { 
        $requiredOUs | ForEach-Object {
            $schoolName = $($validschools.$([int]$PSItem.'School_id'))
            $ou = "ou=Students,ou=$($schoolName),$($domain)"
            $ouPath = $ou.split(',')[1..($ou.Length-1)] -join(',')
            $groupName = "Student Management Accounts - $($schoolName)"
            if (-Not(Get-AdGroup -Filter { SamAccountName -eq $groupName } -ErrorAction SilentlyContinue)) {
                if ($staging) {
                    Write-Host "Staging: New Security Group `"$($groupName)`" at $($ouPath)" -ForeGroundColor Yellow
                    New-ADGroup -Name $groupName -Path $ouPath -GroupScope Global -WhatIf
                } else {
                    New-ADGroup -Name $groupName -Path $ouPath -GroupScope Global
                }
            }
            try {
                Write-Host "Setting permissions for `"$groupName`" on $($ou)"
                if ($staging) {
                    Write-Host "Staging: Grant Password Reset for `"$($groupName)`" on $($ou)" -ForeGroundColor Yellow
                } else {
                    Grant-PasswordResetOnOU -group $groupName -ou $ou
                }
            } catch {
                Write-Host "Error: Failed to set permissions for the Management Account `"$groupName`" on $($ou)."
                Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Failed to set permissions for the Management Account `"$groupName`" on $($ou)."
                exit(1)
            }
        }
    }
    4 { 
        #permissions must be set on each grade
        $ou = "ou=Students,$($domain)"
        $requiredOUs | ForEach-Object {
            $groupName = "Student Management Accounts - $($validschools.$([int]$PSItem.'School_id'))"
            if (-Not(Get-AdGroup -Filter { SamAccountName -eq $groupName } -ErrorAction SilentlyContinue)) {
                if ($staging) {
                    Write-Host "Staging: New Security Group `"$($groupName)`" at $($ou)" -ForeGroundColor Yellow
                    New-ADGroup -Name $groupName -Path $ou -GroupScope Global -WhatIf
                } else {
                    New-ADGroup -Name $groupName -Path $ou -GroupScope Global
                }
            }
            $grade = $PSItem.'Grade'

            switch ($grade) {
                'Prekindergarten' { $grade = 'PK' }
                'Kindergarten' { $grade = 'K' }
            }

            $gradeOU = "ou=$($grade),ou=Students,$($domain)"

            try {
                if ($staging) {
                    Write-Host "Staging: Grant Password Reset for `"$($groupName)`" on $($ou)" -ForeGroundColor Yellow
                } else {
                    Write-Host "Notify: Setting permissions for `"$groupName`" on $($gradeOU)"
                    Grant-PasswordResetOnOU -group $groupName -ou $gradeOU
                }
            } catch {
                Write-Host "Error: Failed to set permissions for the Management Account `"$groupName`" on $($gradeOU)."
                Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Failed to set permissions for the Management Account `"$groupName`" on $($gradeOU)."
                exit(1)
            }
        }
    }
}

#####################################################################
# Home Directory Path Check and Permissions for the Management Groups
#####################################################################
if (@(1,2,3) -contains $adStructure) {
    $reqHomeDir = $studentsCSV | Select-Object -Property @{Name='School_id';Expression={[int]$PSItem.'School_id'}},@{Name='Folder_name';Expression={[int]$PSItem.'Student_gradyr'}} -Unique
} elseif (@(3,4,5) -contains $adStructure) {
    $reqHomeDir = $studentsCSV | Select-Object -Property @{Name='School_id';Expression={[int]$PSItem.'School_id'}},@{Name='Folder_name';Expression={$($PSItem.'Grade').Trim()}} -Unique
}

$validschools.Keys | ForEach-Object {
    $schoolId = $PSItem
    $schoolShortName = $validschools.$PSItem
    $groupName = "Student Management Accounts - $($schoolShortName)"
    #if defined then we need to create the base folders for the gradyrs and assign permissions to the management group.
    if ($homeDirectoryRoot.$schoolId) {
        $homeDirRoot = $homeDirectoryRoot.$schoolId + $schoolShortName
        Write-Host "Testing path: $($homeDirRoot)"

        #if the folder doesn't exist attempt to create it. Then we will test if it exists next.
        if (-not(Test-Path $homeDirRoot)) {
            if ($staging) { 
                Write-Host "Staging: Create home directory root folder $homeDirRoot" -ForeGroundColor Yellow
                New-Item $homeDirRoot -ItemType Directory -WhatIf
            } else {
                $newFolder = New-Item $homeDirRoot -ItemType Directory -ErrorAction SilentlyContinue
            }
        }

        #test-path then find all years and assign permissions
        if (Test-Path $homeDirRoot) {
            #We need to create the folders needed on GRADYR or GRADE from $reqHomeDir. If we are having to create folders then chance is we need to modify ntfs perms for the management group at the root.
            $reqHomeDir | Where-Object { $PSItem.School_id -eq $schoolId } | Select-Object -ExpandProperty Folder_name | ForEach-Object {

                #if using grades then shorten to PK and K.
                if ($PSItem -eq 'Prekindergarten') { $PSItem = 'PK' }
                if ($PSItem -eq 'Kindergarten') { $PSItem = 'K' }

                $folder = "$($homeDirRoot)\$($PSItem)"
                if (-Not(Test-Path $folder)) {
                    if ($staging) { 
                        Write-Host "Staging: Creating Student home directory folder $($folder)" -ForeGroundColor Yellow
                    } else {
                        try {
                            New-Item "$($homeDirRoot)\$($PSItem)" -type Directory -Force
                            if ($(Get-NTFSInheritance $homeDirRoot).AccessInheritanceEnabled) {
                                Write-Host "WARNING: NTFS Inhereitance is enabled at the root of the home directory for $($homeDirRoot). This could lead to leaking data. If this was done on purpose then ignore this message."
                            }
                            if (-Not(Get-NTFSAccess $homeDirRoot | Where-Object { $PSItem.'Account' -like "*\$($groupName)"})) {
                                Add-NTFSAccess -Path $homeDirRoot -Account $groupName -AccessRights $stuManagementHomeDirPerms
                            }
                        } catch {
                            Write-Host "Error: Can not create graduation year folders in $($homeDirRoot). Please check NTFS permissions." -ForegroundColor Red
                            Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Can not create graduation year folders in $($homeDirRoot). Please check NTFS permissions."
                            exit(1)
                        }
                    }
                }
            }
        } else {
            if ($staging) { 
                Write-Host "Staging: Create home directory root folder $homeDirRoot" -ForeGroundColor Yellow
            } else {
                Write-Host "Error: Home Directory path for building $($PSItem) at $($homeDirectoryRoot.$([int]$PSItem)) can not be found. Please fix before running script again." -ForegroundColor Red
                Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Home Directory path for building $($PSItem) at $($homeDirectoryRoot.$([int]$PSItem)) can not be found. Please fix before running script again."
                exit(1)
            }
        }
    }

}

#####################################################################
# UPN Suffixes from $stuEmailDomain
#####################################################################

$upnSuffixes = Get-ADforest $(Get-ADDomain).dnsroot | Select-Object -ExpandProperty UPNSuffixes
$stuEmailDomain.Values | ForEach-Object {
    $upnSuffix = $PSItem
    if ($upnSuffix[0] -eq '@') {
        $upnSuffix = $upnSuffix.Substring(1,$($upnSuffix).Length-1)
    }
    if ($upnSuffixes -notcontains $upnSuffix) {
        if ($staging) {
            Write-Host "Staging: The UPN Suffix $upnSuffix is not known by the domain" -ForeGroundColor Yellow
            Set-ADForest $(Get-ADDomain).dnsroot -UPNSuffixes @{Add=$upnSuffix} -WhatIf
        } else {
            Write-Host "Info: The UPN Suffix $upnSuffix is not known by the domain. Creating..."
            Set-ADForest $(Get-ADDomain).dnsroot -UPNSuffixes @{Add=$upnSuffix}
        }
    }
}

#####################################################################
# Student Accounts
#####################################################################


$newStudents = @() #In order to send an email at the end for all new students.
$errorCount = 0
$errorMessage = @()

if (-Not($SkipStudents)) { #skip this whole process.

#Unfortunately this can not be a Parallel ForEach loop. It locks the DCs.
$studentsCSV | ForEach-Object {

    $student = $PSItem
    $studentId = [int]$student.'Student_id'

    #excluded accounts shouldn't have anything done to them.
    if ($excludedStudentIDs -contains $studentId) { return }

    $givenName = $student.'First_name' #This is done for a reason. Can't remember why.
    $surName = $student.'Last_name' #This is done for a reason. Can't remember why.

    $firstName = Remove-Spaces(Remove-SpecialCharacters($student.'First_name'))

    #if $useNickname in settings.ps1
    if ($useNickname) {
        if ($($student.'Student_nickname').length -gt 1) { $firstName = Remove-Spaces(Remove-SpecialCharacters($student.'Student_nickname')) }
        if ($($student.'Student_nickname').length -gt 1) { $givenName = $student.'Student_nickname' }
    }

    $middleInitial = $student.'Middle_name'[0]
    $lastName = Remove-SpecialCharacters($student.'Last_name') #remove spaces later.

    #if $useFirstHypenatedLastName in settings.ps1
    if ($useFirstHypenatedLastName) {
        if ($lastName.indexOf('-') -gt 1) {
            $lastName = $lastName.Split('-')[0]
            $surName = $surName.Split('-')[0]
        }
    }
    
    #if $useFirstSpacedLastName in settings.ps1
    if ($useFirstSpacedLastName) {
        if ($lastName.indexOf(' ') -gt 1) {
            $lastName = $lastName.Split(' ')[0]
            $surName = $surName.Split(' ')[0]
        }
    }

    $lastName = Remove-Spaces($lastName) #Remove spaces after evaulating if $useFirstSpacedLastName is $True.

    switch ($student.'Grade') {
        'Prekindergarten' { $grade = 'PK' }
        'Kindergarten' { $grade = 'K' }
        'KG' { $grade = 'K' }
        'KF' { $grade = 'K' }
        default { $grade = $student.'Grade' }
    }

    #Overrides
    if ($overridedAccounts.$studentId) {
        $override = $overridedAccounts.$studentId
        $firstName = Remove-Spaces(Remove-SpecialCharacters($override.First_name))
        $givenName = $firstName
        $lastName = Remove-Spaces(Remove-SpecialCharacters($override.Last_name))
        $surName = $lastName
        $middleInitial = Remove-Spaces(Remove-SpecialCharacters($override.MiddleInitial))[0]
        #$grade = $override.Grade
        Write-Host "Info: Student $($override.Student_id) will use $firstName $lastName per the overrides.csv file."
    }

    $fullName = $firstName + ' ' + $lastName

    $gender = $student.'Gender'
    $buildingNumber = [int]$student.'School_id'
    $buildingShortName = $validschools.$([int]$student.'School_id')
    $gradyr = $student.'Student_gradyr'

    if (($null -eq $gradyr) -or ($gradyr -eq '')) {
        Write-Host "Error: Missing GradYR for student $($studentId). Skipping Student as we need that data."
        $errorCount++
        $errorMessage += @("Error: Missing GradYR for student $($studentId). Skipping Student as we need that data.")
        return
    }

    if ($stuEmailDomain.$([int]$student.'School_id')) {
        $emailDomain = $stuEmailDomain.$([int]$student.'School_id')
    } else {
        $emailDomain = $stuEmailDomain.'default'
    }

    if ($RandomPassword) {
        $password = [string]$(Get-RandomCharacters 8 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNOPQRSTUVWXYZ123456789!.$#%&*<>') #no l,o, or 0. So there is no confusion.
    } else {
        $password = Get-NewPassword -student $student
    }
    
    #samAccountName has to 20 characters or less.
    #principalName can be their entire name regardless but at login they must type the entire UPN.

    switch ($stuTemplate) {
        1 { #[firstname][lastname]
            $username = "$($firstName)$($lastName)"
            $principalName = $username + $emailDomain
            $emailAddress = $username + $emailDomain
            if ($username.Length -gt 20) { $username = $username.Substring(0,20) }
        }
        2 { #[firstinitial].[lastname].[gradyr]
            $username = "$($firstName)$($lastName)"
            $principalName = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            $emailAddress = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            if ($username.Length -gt 18) { $username = $username.Substring(0,18) }
            $username = $username + $([string]$gradyr).Substring(2,2)
        }
        3 { #[firstname].[lastname]
            $username = "$($firstName + '.' + $lastName)"
            $principalName = $username + $emailDomain
            $emailAddress = $username + $emailDomain
            if ($username.Length -gt 20) { $username = $username.Substring(0,20) }
        }
        4 { #[firstname].[lastname][gradyr]
            $username = "$($firstName + '.' + $lastName)"
            $principalName = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            $emailAddress = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            if ($username.Length -gt 18) { $username = $username.Substring(0,18) }
            $username = $username + $([string]$gradyr).Substring(2,2) 
        }
        5 { #[lastname].[firstname]
            $username = "$($lastName + '.' + $firstName)"
            $principalName = $username + $emailDomain
            $emailAddress = $username + $emailDomain
            if ($username.Length -gt 20) { $username = $username.Substring(0,20) }
        }
        6 { #[lastname].[firstintial].[gradyr]
            $username = "$($lastName + '.' + $firstName[0])"
            $principalName = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            $emailAddress = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            if ($username.Length -gt 17) { $username = $username.Substring(0,17) }
            $username = $username + '.' + $([string]$gradyr).Substring(2,2) 
        }
        7 { #[firstname].[first3oflastname][last2ofgradyr]
            if ($lastName.length -gt 3) { $lnUsername = ($lastName.Substring(0,3)) } else { $lnUsername = $lastName }
            $username = "$($firstName + '.' + $lnUsername)"
            $principalName = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            $emailAddress = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            if ($username.Length -gt 18) { $username = "$($($firstName.Substring(0,14)) + '.' + $($lnUsername))" } #shortens first name to meet 20 characters
            $username = $username + $([string]$gradyr).Substring(2,2)
        }
        8 { #[first5oflastname][first5offirstname][last2ofgradyr] #This will never be longer than the maximum SAMAccountName length.
            if ($firstName.length -gt 5) { $fnUsername = ($firstName.Substring(0,5)) } else { $fnUsername = $firstName }
            if ($lastName.length -gt 5) { $lnUsername = ($lastName.Substring(0,5)) } else { $lnUsername = $lastName }
            $username = "$($lnUsername)$($fnUsername)" + $([string]$gradyr).Substring(2,2)
            $principalName = $username + $emailDomain
            $emailAddress = $username + $emailDomain
        }
    }

    if ($($homeDirectoryRoot.$buildingNumber)) {
        $homeDirRoot = $homeDirectoryRoot.$buildingNumber + $($buildingShortName) + '\'
        if (@(1,2,3) -contains $adstructure) {
            $homeDirPath = "$($homeDirRoot)$($gradyr)"
        } elseif (@(4,5,6) -contains $adStructure) {
            $homeDirPath = "$($homeDirRoot)$($grade)"
        }
        $homeDir = "$($homeDirPath)\$($username)"
    } else {
        #clear for the next loop
        $homeDirRoot = $null; $homeDirPath = $null; $homeDir = $null
    }

    #expected OU to create user in
    switch ($adStructure) {
        1 { $ou = "ou=$($gradyr),ou=Students,$($domain)" }
        2 { $ou = "ou=$($gradyr),ou=$($buildingShortName),ou=Students,$($domain)" }
        3 { $ou = "ou=$($gradyr),ou=Students,ou=$($buildingShortName),$($domain)" }
        4 { $ou = "ou=$($grade),ou=Students,$($domain)" }
        5 { $ou = "ou=$($grade),ou=$($buildingShortName),ou=Students,$($domain)" }
        6 { $ou = "ou=$($grade),ou=Students,ou=$($buildingShortName),$($domain)" }
    }

    #We have to check if user exists in AD already even if they are disabled based on their Student ID.
    #We need to search the entire directory as its possible a student is moving between buildings.
    #Need to try catch the errors. If found means the try was succesful. If not found we move on.
    #if the error is unrecoverable we need to break from creating accounts.
    try {
        $existingAccount = Get-AdUser -Filter { EmployeeNumber -eq $studentId } -Properties *
    # } catch [Microsoft.ActiveDirectory.Management.ADException] {
    #     Write-Host "Error: ADException trying to query $studentId", $Error[0].Exception.GetType().fullname
    #     $errorFullName = $Error[0].Exception.GetType().fullname
    #     Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Different type of error query for $($studentId). $($errorFullName)"
    #     break
    } catch {
        $errorFullName = $Error[0].Exception.GetType().fullname
        Write-Host "Error: Error while querying domain for student $($studentId). Breaking from processing students. The error was:", $Error[0].Exception.GetType().fullname
        Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Error while querying domain for student $($studentId). $($errorFullName)"
        break
    }

    if ($existingAccount) {
        #user account already exist.
        
        if ($SkipExistingStudents) { return }

        $existingAccountGUID = $existingAccount.ObjectGUID
        $oldName = $existingAccount.name

        #There is a lot that goes on for an existing account. If staging evaluate some info here.
        if ($staging) { 
        
            if (Test-Path $currentPath\x_InterimProcessingExistingAccounts.ps1) {
                . $currentPath\x_InterimProcessingExistingAccounts.ps1
            }
        
            if ($VerboseStudent) { Write-Host "Staging: Account exists for $username, $fullname, $gradyr, $buildingShortName, at $($existingAccount.'DistinguishedName')" -ForeGroundColor Yellow }

            if (($existingAccount.GivenName -ne "$givenName") -or ($existingAccount.surname -ne "$surName")) {
                Write-Host "Staging: $($studentId) This account would experience a name change detection. Would change from ""$($existingAccount.GivenName) $($existingAccount.surname)"" to ""$givenName $surName""" -ForegroundColor Green
            }

            if (($existingAccount.Mail -ne "$emailAddress") -or ($existingAccount.userPrincipalName -ne "$emailAddress")) {
                Write-Host "Staging: $($studentId) Email and UserPrincipalName MIGHT change from $($existingAccount.userPrincipalName) to $($emailAddress) if it is available." -ForegroundColor Green

            }

            if ("$($existingAccount.homeDirectory)" -ne "$homeDir") {
                Write-Host "Staging: $($studentId) Home Directory Mismatch. This would change the home directory from ""$($existingAccount.homeDirectory)"" to ""$homeDir""" -ForegroundColor Green
            }

            $existingOU = $($($existingAccount.DistinguishedName).split(',')[1..($($existingAccount.DistinguishedName).Length -1 )] -join ',')
            if ($existingOU -ne $ou) {
                Write-Host "Staging: $($studentId) This account would move from $existingOU to $ou" -ForegroundColor Green
            }
            
            if ($stuGradYrRename){
                if ($existingAccount.Fax -ne $gradyr) {
                    if (@(2,4,6,7,8) -contains $stuTemplate) {
                        Write-Host "Staging: $($studentId) This account has a different graduation year than before. This account will be renamed to reflect the graduation year change." -ForegroundColor Green
                    }
                }
            }
            return
        }

        if ($VerboseStudent) { Write-Host "Verbose: Account already exists for $($fullname)." -ForegroundColor Yellow }

        #Enable Account if Disabled (returning student)
	    if ($existingAccount.Enabled -eq $False) {
            try {
                Write-Host "Info: Enabling account $($existingAccount.'DistinguishedName')" -ForeGroundColor Yellow
                Set-AdAccountPassword -Identity $existingAccountGUID -Reset -NewPassword (ConvertTo-SecureString "$password" -AsPlainText -Force)
                Invoke-SqlUpdate -Query "REPLACE INTO passwords (Student_id, Student_password, HAC_passwordset, Timestamp) VALUES ($studentid,""$password"",NULL,$timestamp)" | Out-Null
                Set-AdUser -Identity $existingAccountGUID -ChangePasswordAtLogon $False
                Set-ADAccountControl -Identity $existingAccountGUID -CannotChangePassword $False
                Enable-ADAccount -Identity $existingAccountGUID

                #Log Event Enable
                Invoke-SqlUpdate -Query "INSERT INTO action_log (Student_id, Identity, Action, Timestamp) VALUES ($studentid,""$($existingAccount.UserPrincipalName)"",""Enable"",""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"")" | Out-Null

                if ($sendMailNotifications) {

                    if ($notifyEmails.$buildingNumber) {
                        $mailTo = $notifyEmails.$buildingNumber
                    } else {
                        $mailTo = @("$sendMailToEmail" )
                    }

                    Write-Host "Info: Sending notification email to $($mailto -join ',') about $emailAddress's enabled account."
                    Send-EmailNotification -mailto $mailTo -subject "Notice: Enabled Student Account for $emailAddress" -body (((((($existingAccountMessage -replace "{USEREMAILADDRESS}",$emailAddress) -replace "{PASSWORD}",$password) -replace "{FULLNAME}",$fullName) -replace "{GRADE}",$grade) -replace "{SCHOOLSHORTNAME}",$buildingShortName) -replace "{STUDENTID}",$studentId)
                    
                }

                #We need this account info in the post processing script for any tasks.
                $newStudents += [PSCustomObject]@{
                    Student_id = $studentid
                    School_id = $buildingNumber
                    Building = $buildingShortName
                    Name = $fullname
                    Grade = $grade
                    EmailAddress = $emailAddress
                    Username = $username
                    Password = $password
                }

            } catch {
                Write-Host "Error: Problem enabling $($existingAccount.distinguishedname)." -ForegroundColor Red
                $errorCount++
                $errorMessage += "Error: Problem enabling account $($existingAccount.distinguishedname)."
                return #keep going. This isn't the end of the world.
            }
        }
        
        #If you use a student graduation year as part of their username, and their grade changes, do you want the account renamed?
        if ($stuGradYrRename -and (-Not($DisableRenamingAccounts))) {
            #if ($existingAccount.Fax -ne $gradyr) { #If a student is returning and their gradyr wasn't in the .Fax field it caused a failure. So we need to actually check the username.
            if ( ($($existingAccount.UserPrincipalName).Substring($($($existingAccount.UserPrincipalName).Split('@')[0]).Length - 2,2)) -ne $([string]$gradyr).Substring(2,2) ) {
                if (@(2,4,6,7,8) -contains $stuTemplate) {
                    $existingAccount.surname = '' #Changing the name will force a rename in the next code block which should fix the gradyr change.
                }
            }
        }

        #Rename Account if name mismatch. Only on Firstname and Given name as username may change on duplicate.
        if ((($existingAccount.GivenName -ne "$givenName")`
         -or ($existingAccount.surname -ne "$surName")) -and (-Not($DisableRenamingAccounts))) {
            Write-Host "Notify: Name change detected. Updating $($existingAccount.'samaccountname') to $($username)"

            #Test if new account name is available.
            if ($(Get-AdUser -Filter "(ObjectGUID -ne ""$existingAccountGUID"") -and ((SamAccountName -eq ""$($username)"") -or (UserPrincipalName -eq ""$($principalName)""))")) {
                if ($newUsername = Get-NextAvailableUsername $username $principalName $homeDirPath $firstName $lastName $middleInitial) {
                    #$newUsername
                    Write-Host "Notify: The username $($username) was already taken. Using $($newUsername.'username') instead."
                    $username = $newUsername.'username'
                    $principalName = $newUsername.'principalName'
                    $emailAddress = $newUsername.'principalName'
                    
                    if ($($homeDirectoryRoot.$buildingNumber)) {
                        $homeDir = $newUsername.'homedir'
                    }
                } else {
                    Write-Host "Error: Could not get an available username for $username. You should never see this message unless you have a serious issue with username conflicts. You need to reconsider your student usernames." -ForeGroundColor Red
                    Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Could not get an available username for $username. You should never see this message unless you have a serious issue with username conflicts. You need to reconsider your student usernames."
                    exit(1)
                }
            }

            try {
                Set-ADUser -Identity $existingAccountGUID `
                -SamAccountName $username `
                -givenName $givenName `
                -Surname $surName `
                -DisplayName $fullName `
                -UserPrincipalName $principalName `
                -EmailAddress $emailAddress `

                Rename-ADObject -Identity $existingAccountGUID -NewName $principalName

                #Log Event Rename
                Invoke-SqlUpdate -Query "INSERT INTO action_log (Student_id, Identity, Action, Timestamp) VALUES ($studentid,""$($existingAccount.SamAccountName) to $principalName"",""Rename"",""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"")" | Out-Null

                #move home directory
                if ($homedir) {
                    if (Test-Path "$($existingAccount.homeDirectory)") {
                        if ($VerboseStudent) {  Write-Host "Comparing $($existingAccount.homeDirectory) to $homeDir" }
                        if ("$($existingAccount.homeDirectory)" -ne "$homeDir") {
                            #Move-item doesn't move between servers. So robocopy with move then rename old folder as robocopy doesn't always remove them and file locks can be an issue.
                            Start-Process -FilePath robocopy.exe -ArgumentList """$($existingAccount.'homeDirectory')"" ""$($homeDir)"" /mir /mov /copy:dats /nfl /ndl /np /njh /w:30 /r:1" -NoNewWindow -Wait 
                            #We never want to lose data. Rename instead of Remove. Should be obvious what folders need cleaned up later. File locks can mess this process up.
                            try {
                                Rename-Item "$($existingAccount.'homeDirectory')" "$([string]$(Get-Date -UFormat %s -Millisecond 0))-$($username)" -Force
                            } catch {
                                Write-Host "Warning: Problem moving old home directory folder $($existingAccount.'homeDirectory') for $($username). New path is already set and this shouldn't cause a problem." -ForegroundColor Yellow
                            }
                            Set-ADUser -Identity $existingAccountGUID -HomeDrive 'h:' -HomeDirectory $homeDir
                        }
                    } else {
                        Set-ADUser -Identity $existingAccountGUID -HomeDrive 'h:' -HomeDirectory $homeDir
                        New-HomeDirectory $username $homeDir
                    }
                }

                if ($sendMailNotifications) {

                    if ($notifyEmails.$buildingNumber) {
                        $mailTo = $notifyEmails.$buildingNumber
                    } else {
                        $mailTo = @("$sendMailToEmail" )
                    }

                    #$mailTo = $notifyEmails.$buildingNumber
                    Send-EmailNotification -mailto $mailTo -subject "Notice: Name Change for $($emailAddress)" -body (((((($nameChangeMessage -replace "{USEREMAILADDRESS}",$emailAddress) -replace "{OLDUSEREMAILADDRESS}",$existingAccount.Mail) -replace "{FULLNAME}",$fullName) -replace "{GRADE}",$grade) -replace "{SCHOOLSHORTNAME}",$buildingShortName) -replace "{STUDENTID}",$studentId)
                }

            } catch {
                Write-Host "Error: Failed to rename $oldName to $username. $PSItem"
                $errorCount++
                $errorMessage += @("Error: Failed to rename $oldName to $username")
                return
            }            
        } else {
            #If the home directory field was null but now we need to create one and update the user account.
            if ($homedir -and ($existingAccount.homeDirectory -eq $null)) {
                New-HomeDirectory $existingAccount.SamAccountName $homeDir
                Set-ADUser -Identity $existingAccountGUID -HomeDrive 'h:' -HomeDirectory $homeDir
            }
        }

        
        $existingOU = $($($existingAccount.DistinguishedName).split(',')[1..($($existingAccount.DistinguishedName).Length -1 )] -join ',')

        #Check if account is in any of the special OUs that we don't move students out of.
        if ($specialOUs) {
            $specialOUs | ForEach-Object {
                if ($existingOU -like "*OU=$($PSItem)*") {
                    #if current Restricted OU is not like the calculated OU then we have to assume a building change.
                    if ($ou -notlike "*$(($existingAccount.DistinguishedName).split(',')[(($($existingAccount.DistinguishedName).Split(',').indexof('OU=$($PSItem)'))+1)..($($existingAccount.DistinguishedName).Split(',').Length)] -join ',')") {
                        #Set $ou to the new building but keep it in the restricted OU for the new building.
                        $ou = "OU=$($PSItem),$($($ou).split(',')[1..($($ou).Length -1 )] -join ',')"
                    } else {
                        $ou = $existingOU
                        #Write-Host "Info: Not moving student $($existingAccount.DistinguishedName) from $($PSItem) OU."
                    }
                }
            }
        }

        #Update grade, building, and Student_gradyr information. Move OU and HomeDir if needed. Think graduating between buildings.
        
        if (($existingOU -ne $ou)`
        -or ($existingAccount.Fax -ne $gradyr)`
        -or ($existingAccount.homePhone -ne $grade)`
        -or ($existingAccount.physicalDeliveryOfficeName -ne $buildingShortName)`
        ) {
            try {
                Set-ADUser -Identity $existingAccountGUID -Fax $gradyr -Office $buildingShortName -HomePhone $grade
                
                if ($existingOU -ne $ou) {
                    Write-Host "Info: Moving $existingAccount to $ou"
                    Move-ADObject -Identity $existingAccountGUID -TargetPath $ou
                } 

                if ($homedir) {
                    if (Test-Path "$($existingAccount.homeDirectory)") {
                        if ($VerboseStudent) {  Write-Host "Comparing $($existingAccount.homeDirectory) to $homeDir" }
                        if ("$($existingAccount.homeDirectory)" -ne "$homeDir") {
                            #Move-item doesn't move between servers. So robocopy with move then rename old folder as robocopy doesn't always remove them and file locks can be an issue.
                            Start-Process -FilePath robocopy.exe -ArgumentList """$($existingAccount.'homeDirectory')"" ""$($homeDir)"" /mir /mov /copy:dats /nfl /ndl /np /njh /w:30 /r:1" -NoNewWindow -Wait 
                            #We never want to lose data. Rename instead of Remove. Should be obvious what folders need cleaned up later. File locks can mess this process up.
                            try {
                                Rename-Item "$($existingAccount.'homeDirectory')" "$($username)-$([string]$(Get-Date -UFormat %s -Millisecond 0))" -Force
                            } catch {
                                Write-Host "Warning: Problem moving old home directory folder $($existingAccount.'homeDirectory') for $($username). New path is already set and this shouldn't cause a problem." -ForegroundColor Yellow
                            }
                            Set-ADUser -Identity $existingAccountGUID -HomeDrive 'h:' -HomeDirectory $homeDir
                        }
                    } else {
                        Set-ADUser -Identity $existingAccountGUID -HomeDrive 'h:' -HomeDirectory $homeDir
                        New-HomeDirectory $username $homeDir
                    }
                }

            } catch {
                Write-Host "Error: Failed to update and move $username to $ou" -ForegroundColor Red
                $errorCount++
                $errorMessage += @("Error: Failed to update and move $username to $ou")
                return
            }
        } else {
            #If the home directory field was null but now we need to create one and update the user account.
            if ($homedir -and ($existingAccount.homeDirectory -eq $null)) {
                New-HomeDirectory $existingAccount.SamAccountName $homeDir
                Set-ADUser -Identity $existingAccountGUID -HomeDrive 'h:' -HomeDirectory $homeDir
            }
        }

        #blank home directory if something has changed to make it null
        if (-Not($homeDir)) {
            #blank home directory field if the building doesn't have one specified.
            Set-ADUser -Identity $existingAccountGUID -HomeDrive $null -HomeDirectory $null
        }

        #This script is called and has access to all variables of $student to customize for individual schools.
        if (Test-Path $currentPath\x_InterimProcessingExistingAccounts.ps1) {
            . $currentPath\x_InterimProcessingExistingAccounts.ps1
        }

        $homeDirRoot = $null; $homeDirPath = $null; $homeDir = $null

        if ($StopAfterXExisting -ge 1) {
            $existingAccountsCount++
            if ($existingAccountsCount -ge $StopAfterXExisting) {
                Write-Host "Info: Threshold of $StopAfterXExisting on processing existing accounts has been met. Breaking from modifying students."
                break
            }
        }
        
    } else {
        #we need to create a new account for the student. If we have a conflict we need to decide how to cope with it.
        #if ($staging) { Write-Host "Staging: Create an account for $username, $fullname, $gradyr, $buildingShortName, at $ou" -ForeGroundColor Yellow }

        if ($SkipNewStudents) { return }
        if ($VerboseStudent) { Write-Host "Info: Need to verify we can create a new account for $username, $fullname, $gradyr, $buildingShortName, at $ou" -ForeGroundColor Yellow }

        #need to find out if the username is already taken by another user.
        if ($(Get-AdUser -Filter "(SamAccountName -eq ""$($username)"") -or (UserPrincipalName -eq ""$($principalName)"")")) {
            Write-Host "Error: The username for $username ($studentid) is already taken." -ForegroundColor Red
            #we need to return a hash table containing the new username,userprincipalname, and homedir.
            #attempt to create the new user with the middle initial while matching the $adStructure.

            $existingAccountGUID = 'blank' #must be a non empty string

            if ($newUsername = Get-NextAvailableUsername $username $principalName $homeDirPath $firstName $lastName $middleInitial) {
                #$newUsername
                $username = $newUsername.'username'
                $principalName = $newUsername.'principalName'
                $emailAddress = $newUsername.'principalName'

                if ($($homeDirectoryRoot.$buildingNumber)) {
                    $homeDir = $newUsername.'homedir'
                }

            } else {
                Write-Host "Error: Could not get an available username for $username. You should never see this message unless you have a serious issue with username conflicts. You need to reconsider your student usernames." -ForeGroundColor Red
                    Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Could not get an available username for $username. You should never see this message unless you have a serious issue with username conflicts. You need to reconsider your student usernames."
                exit(1)
            }
        }

        Write-Host "Info: Create an account for $username, $fullname, $studentid, $gradyr, $buildingShortName, $password, at $ou" -ForeGroundColor Yellow

        Write-Host $homeDir
        if ($homeDir) {
            if (-Not(Test-Path $homeDirRoot)) {
                Write-Host "Error: Unable to access the home directory root at $homeDirRoot. Please fix in settings.ps1 for this building. Not creating student account until resolved." -ForegroundColor Red
                if (-Not($staging)) { 
                    $errorCount++
                    $errorMessage += @("Error: Unable to access the home directory root at $homeDirRoot. Please fix in settings.ps1 for this building. Not creating student account until resolved.")
                    exit(1)
                }
            }
        }

        #if you use the FullName for name then you can run into a distinguishedname conflict.
        try {

            if ($staging) {
                
                New-Aduser `
                -sAMAccountName $username `
                -givenName $givenName `
                -Surname $surName `
                -DisplayName $fullName `
                -name $principalName `
                -EmployeeNumber $studentId `
                -ChangePasswordAtLogon $false `
                -AccountPassword (ConvertTo-SecureString "$password" -AsPlainText -force) `
                -Enabled $true  `
                -Path $ou `
                -Office $buildingShortName `
                -HomePhone $grade `
                -Fax $gradyr `
                -EmailAddress $emailAddress `
                -UserPrincipalName $principalName `
                -WhatIf
                
                return

            } else {
                
                #We need the email address to exist in Google before we create and set the password otherwise the Google Account will be locked out and miss the GAPS sync.
                if ($GAMprecreateUser) {
                    if (-Not($GAMDefaultOrg)) { $GAMDefaultOrg = '/' }
                    & $currentPath\..\gam\gam.exe create user $emailAddress firstname "$givenName" lastname "$surName" org "$GAMDefaultOrg"
                    if ($LASTEXITCODE -eq 57) {
                        Write-Host "Error: Failed to precreate user $emailAddress in Google G Workspaces due to duplicate."
                        if ($GAMStopOnDuplicate) {
                            $errorCount++
                            $errorMessage += @("Error: Failed to precreate user $emailAddress in Google G Workspaces due to duplicate.")
                            return
                        }
                    } elseif ($LASTEXITCODE -ge 1) {
                        Write-Host "Error: Failed to precreate user $emailAddress in Google G Workspaces."
                        $errorCount++
                        $errorMessage += @("Error: Failed to precreate user $emailAddress in Google G Workspaces.")
                        return
                    }
                }

                New-Aduser `
                -sAMAccountName $username `
                -givenName $givenName `
                -Surname $surName `
                -DisplayName $fullName `
                -name $principalName `
                -EmployeeNumber $studentId `
                -ChangePasswordAtLogon $false `
                -AccountPassword (ConvertTo-SecureString "$password" -AsPlainText -force) `
                -Enabled $true  `
                -Path $ou `
                -Office $buildingShortName `
                -HomePhone $grade `
                -Fax $gradyr `
                -EmailAddress $emailAddress `
                -UserPrincipalName $principalName `

                Invoke-SqlUpdate -Query "REPLACE INTO passwords (Student_id, Student_password, HAC_passwordset, Timestamp) VALUES ($studentid,""$password"",NULL,$timestamp)" | Out-Null
                
                #Log Event Create
                Invoke-SqlUpdate -Query "INSERT INTO action_log (Student_id, Identity, Action, Timestamp) VALUES ($studentid,""$principalName"",""Create"",""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"")" | Out-Null

                if ($sendMailNotifications) {

                    if ($notifyEmails.$buildingNumber) {
                        $mailTo = $notifyEmails.$buildingNumber
                    } else {
                        $mailTo = @("$sendMailToEmail" )
                    }

                    #$mailTo = $notifyEmails.$buildingNumber
                    Send-EmailNotification -mailto $mailTo -subject "Notice: New Student Account for $($emailAddress)" -body (((((($newAccountMessage -replace "{USEREMAILADDRESS}",$emailAddress) -replace "{PASSWORD}",$password) -replace "{FULLNAME}",$fullName) -replace "{GRADE}",$grade) -replace "{SCHOOLSHORTNAME}",$buildingShortName) -replace "{STUDENTID}",$studentId)
                }

                if ($homeDir) {
                    #This does not create the home directory. It only sets the value.
                    Get-ADUser $username | Set-ADUser -HomeDrive 'h:' -HomeDirectory $homeDir
                    #This does the actual foldder creation and permissions.
                    New-HomeDirectory $username $homeDir
                }

                $newStudents += [PSCustomObject]@{
                    Student_id = $studentid
                    School_id = $buildingNumber
                    Building = $buildingShortName
                    Name = $fullname
                    Grade = $grade
                    EmailAddress = $emailAddress
                    Username = $username
                    Password = $password
                }

                #custom actions to new accounts.
                if (Test-Path $currentPath\x_InterimProcessingNewAccounts.ps1) {
                    . $currentPath\x_InterimProcessingNewAccounts.ps1
                }


            }

            if ($StopAfterXNew -ge 1) {
                $newAccountsCount++
                if ($newAccountsCount -ge $StopAfterXNew) {
                    Write-Host "Info: Threshold of $StopAfterXNew on processing new accounts has been met. Breaking from modifying students."
                    break
                }
            }
            
            $homeDirRoot = $null; $homeDirPath = $null; $homeDir = $null
        } catch {
            Write-Host "Error: Failed to create account for $($fullName), $($username), $($password), at $($ou). Try running script again in case this is a duplicate user error. $PSitem" -ForegroundColor Red
            Write-Host $Error[0].Exception.GetType().fullname
            $homeDirRoot = $null; $homeDirPath = $null; $homeDir = $null
            $errorCount++
            $errorMessage += @("Error: Failed to create account for $($fullName), $($username), at $($ou). Try running script again in case this is a duplicate user error.")
            return
        }
        
    }

}
}

#If there are errors send an email but do not continue with post processing script as it may contain bad information.
if ($errorCount -ge 1) {
    Stop-Transcript
    $errorMessageString = $errorMessage -join "`r`n"
    
    if ($errorCount -gt $maxChanges) {
        Send-EmailNotification -subject "Automated_Students: Failure" -body "There were $errorCount errors detected while processing student accounts.`r`n$($errorMessageString)`r`nPlease inspect the log file $($logfile) for more details.`r`nThis script will not run the post processing script."
        exit(1)
    } else {
        Send-EmailNotification -subject "Automated_Students: Failure" -body "There were $errorCount errors detected while processing student accounts.`r`n$($errorMessageString)`r`nPlease inspect the log file $($logfile) for more details.`r`nThis script will continue as the errors do not exceed the Max Changes in settings.ps1."
    }
}

# Final Script to do all the stuff.
if (-Not($DisablePostProcessingScript)) {
    if (-Not($Staging)) {
        
        #Process Post Processing Scripts
        $PostProcessingScripts = Get-ChildItem -Filter PostProcessingScripts\*.ps1 | Select-Object -ExpandProperty name

        #ASynchronous First
        Write-Host "Info: Starting Post Processing Tasks." -ForegroundColor Yellow
        $PostProcessingScripts | Where-Object { $PSItem -notlike "last_*.ps1" -and $PSItem -notlike "disabled_*.ps1" } | ForEach-Object {
            Write-Host "Info: Running $PSItem" -ForegroundColor Yellow
            Invoke-Expression -Command "& "".\PostProcessingScripts\$PSItem"""
        }

        #Last_*.ps1 Last
        Write-Host "Info: Starting Last Post Processing Tasks." -ForegroundColor Yellow
        $PostProcessingScripts | Where-Object { $PSItem -like "last_*.ps1" } | ForEach-Object {
            Write-Host "Info: Running $PSItem" -ForegroundColor Yellow
            Invoke-Expression -Command "& "".\PostProcessingScripts\$PSItem"""
        }

    } else {
        Write-Host "Info: Staging specified. Will not run Post Processing Scripts."
    }
} else {
    Write-Host "Info: DisablePostProcessingScript specified. Will not run Post Processing Scripts."
}

Stop-Transcript
