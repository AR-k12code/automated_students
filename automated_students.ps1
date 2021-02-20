#Requires -Version 7.0

# Automated Student Account Creation Script
# Craig Millsap, Gentry Public Schools, cmillsap@gentrypioneers.com, 9/2020

# This script was originally designed around the Clever CSV import files.
# It is now based on the SQLite included with this project and is required.

# We should turn on Indexing on the EmployeeNumber attribute

Param(
    [Parameter(Mandatory=$false)][switch]$Staging, #don't make changes to domain
    [Parameter(Mandatory=$false)][int]$Threads=1, #anymore than this seems to cause problems. Probably with a backlog of generating SIDs.
    [Parameter(Mandatory=$false)][switch]$SkipStudents,
    [Parameter(Mandatory=$false)][switch]$SkipNewStudents,
    [Parameter(Mandatory=$false)][switch]$SkipExistingStudents,
    [Parameter(Mandatory=$false)][switch]$ResetAllPasswords,
    [Parameter(Mandatory=$false)][switch]$VerboseStudent,
    [Parameter(Mandatory=$false)][switch]$DisablePostProcessingScript, #Do not run the final script.
    [Parameter(Mandatory=$false)][switch]$RenewEmailPassword,
    [Parameter(Mandatory=$false)][int]$StopAfterXNew, #Testing something on new accounts? Only create this number then quit. Useful while testing x_InterimProcessingNewAccounts.ps1
    [Parameter(Mandatory=$false)][int]$StopAfterXExisting, #Testing on existing accounts? Only create this number then quit. Useful while testing x_InterimProcessingExistingAccounts.ps1
    [Parameter(Mandatory=$false)][switch]$ForceDoNotRequireDomainAdmin, #I'm serious. You can mess things up if you can't read confidentiality bits.
    [Parameter(Mandatory=$false)][switch]$SkipDisablingAccounts #This will skip disabling accounts that have an EmployeeNumber and are not in the exclusions or in the students table. This does nothing for accounts that do not have EmployeeNumbers.
)

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
if (-Not(Test-Path $currentPath\logs)) { New-Item -ItemType Directory -Path $currentPath\logs }
$logfile = "$currentPath\logs\automated_students_$(get-date -f yyyy-MM-dd-HH-mm-ss).log"
try {
    Start-Transcript($logfile)
} catch {
    Stop-TranScript; Start-Transcript($logfile)
}

#Pull in settings
if (Test-Path $currentPath\settings.ps1) {
    . $currentPath\settings.ps1
} else {
    Write-Host "Error: Missing settings.ps1 file. Please read documentation." -ForegroundColor Red
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
# Dependencies - PSCore7, CognosDownloader, ActiveDirectory, NTFSSecurity,PSQLite Modules
#####################################################################

# I attempted to use the new parallel feature but it kept killing my DCs. DO NOT USE THAT FEATURE OF PWSH7.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires Powershell 7 or higher." -foregroundcolor RED
    exit(1)
}


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

#PSSQLite Module is required.
if (Get-Module -ListAvailable | Where-Object {$PSItem.name -eq "PSSQLite"}) {
    Import-Module PSSQLite
  } else {
    Write-Host 'PSSQLite Module not found!'
    if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically install it? y/n')) {
        try {
            Install-Module -Name PSSQLite -Scope AllUsers -Force
            Import-Module PSSQLite
        } catch {
            Write-Host 'Failed to install PSSQLite Module.' -ForegroundColor RED
            exit(1)
        }
    } else {
        exit(1)
    }
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

#Pull in functions
if (Test-Path $currentPath\z_functions.ps1) {
    . $currentPath\z_functions.ps1
} else {
    Write-Host "Error: Missing z_functions.ps1 file. Please read documentation." -ForegroundColor Red
    exit(1)
}

#Active Directory Information
try {
    $domain = $(Get-ADDomain).DistinguishedName
    $domainfqdn = $(Get-ADDomain).Forest
} catch {
    Write-Host 'Error: Unable to Access domain controller.' -ForegroundColor Red
    Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Unable to Access domain controller."
    exit(1)
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
# Required folder for Cognos Reports then loop through and download
# required reports. We are not checking here for failures.
# The CognosDownload.ps1 script does error checking and will send an email notification.
# If downloading updated reports fails the script will continue with existing files.
#####################################################################
try {
    #$studentsCSV = Invoke-SQLiteQuery -DataSource $database -Query "SELECT students.*,Student_nickname,Student_gradyr FROM students LEFT JOIN students_extras ON students.Student_id = students_extras.Student_id ORDER BY Student_id DESC" -ErrorAction 'STOP' 
    #sorting by highest number first causes new accounts to be created first BUT it if you redo things it will cause a younger duplicate name conflict to have precidence.
    $studentsCSV = Invoke-SQLiteQuery -DataSource $database -Query 'SELECT
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
    $adStudents = Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=Students,$domain" -Properties EmployeeNumber,memberof
} elseif (@(3,6) -contains $adStructure) {
    $adStudents = @()
    $stuOUs | ForEach-Object {
        $adStudents += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "$PSItem" -Properties EmployeeNumber,memberof
    }
}

$activeStudentIDs = @()
$activeStudentIDs += $adStudents | Select-Object -ExpandProperty EmployeeNumber

#Calculate changes and if it exceeds our daily maximum then abort.
if (-Not($SkipDisablingAccounts)) {
    #Count New and Deactivated.
    if ((Compare-Object $activeStudentIDs $validStudentIDs -PassThru).count -ge $maxChanges) {
        Write-Host "Error: Calculated changes `($([string]$($(Compare-Object $activeStudentIDs $validStudentIDs -PassThru).count))`) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. Disabling accounts are included in this count."
        if (-Not($staging)) {
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Calculated changes `($([string]$($(Compare-Object $activeStudentIDs $validStudentIDs -PassThru).count))`) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. Disabling accounts are included in this count."
            exit(1)
        }
    } else {
        Write-Host "Info: Current allowed number of new and disabled students is $($maxChanges). This is roughly $([int]($maxChanges / ($studentsCSV | measure-object).Count * 100))% of your students."
    }
} else {
    #Count New Only
    if ((Compare-Object $activeStudentIDs $validStudentIDs | Where-Object { $PSItem.SideIndicator -eq '=>' }).count -ge $maxChanges) {
        Write-Host "Error: Calculated changes `($([string]$($(Compare-Object $activeStudentIDs $validStudentIDs | Where-Object { $PSItem.SideIndicator -eq '=>' }).count))`) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. The disabling of accounts are NOT included in this count."
        if (-Not($staging)) {
            Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Calculated changes `($([string]$($(Compare-Object $activeStudentIDs $validStudentIDs -PassThru).count))`) exceed the set maximum changes of $($maxChanges). Please make sure you adjust your settings.ps1 file. The disabling of accounts are NOT included in this count."
            exit(1)
        }
    } else {
        Write-Host "Info: Current allowed number of new and disabled students is $($maxChanges). This is roughly $([int]($maxChanges / ($studentsCSV | measure-object).Count * 100))% of your students."
    }
}

#Exclusion List
if (Test-Path $currentPath\exclusions.csv) {
    Write-Host "Info: Exclusion CSV detected." -ForegroundColor YELLOW
    $excludeCSV = Import-CSV $currentPath\exclusions.csv
    #$headers = 
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
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Student_gradyr -Unique
        $stuOUs += @("ou=Students,$($domain)","ou=Disabled,ou=Students,$($domain)","ou=Restricted,ou=Students,$($domain)")
        $requiredOUs | ForEach-Object {
            $stuOUs += @("ou=$($PSItem.'Student_gradyr'),ou=Students,$($domain)")
        }
    }
    2 { #2 = STUDENTS/SCHOOL/GRADYR
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Student_gradyr -Unique
        $stuOUs += @("ou=Students,$($domain)","ou=Disabled,ou=Students,$($domain)")
        $validschools.Values | ForEach-Object {
            $stuOUs += @("ou=$($PSItem),ou=Students,$($domain)","ou=Restricted,ou=$($PSItem),ou=Students,$($domain)")
        }
        $requiredOUs | ForEach-Object {
            $stuOUs += @("ou=$($PSItem.'Student_gradyr'),ou=$($validschools.$([int]$PSItem.'School_id')),ou=Students,$($domain)")
        }
    }
    3 { #3 = SCHOOL/STUDENTS/GRADYR
        $requiredOUs = $studentsCSV | Select-Object -Property School_id,Student_gradyr -Unique
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
    2 {
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
    3 { 
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
    5 {
        #permissions need to be set on the school OU only.
        $validschools.Values | ForEach-Object {
            $grade = $PSItem.'Grade'

            switch ($grade) {
                'Prekindergarten' { $grade = 'PK' }
                'Kindergarten' { $grade = 'K' }
            }

            $ou = "ou=$($grade),ou=Students,$($domain)"
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
    6 { 
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
# Password CSV Files Reset
#####################################################################

if (-Not(Test-Path $currentPath\passwords)) { New-Item -ItemType Directory -Path $currentPath\passwords }
if ($ResetAllPasswords) {
    $passwordheader = "Student ID,Full Name,Email Address,Password`r`n"
    $validschools.Values | ForEach-Object {
        $passwordheader | Out-File "$($currentPath)\passwords\$($PSItem)-passwords.csv" -Force -NoNewline -Encoding UTF8
    }
    
    $passwordhashtable = @{}
    $validschools.Values | ForEach-Object {
        $passwordhashtable."$PSItem" = ''
    } 
} else {
    $passwordheader = "Student ID,Full Name,Email Address,Password`r`n"
    $validschools.Values | ForEach-Object {
        if (-Not(Test-Path "$($currentPath)\passwords\$($PSItem)-passwords.csv")) {
            $passwordheader | Out-File "$($currentPath)\passwords\$($PSItem)-passwords.csv" -Force -NoNewline -Encoding UTF8
        }
    }
}

#####################################################################
# Student Accounts
#####################################################################



if (-Not($SkipStudents)) { #skip this whole process.

$newStudents = @() #In order to send an email at the end for all new students.
$errorCount = 0
$errorMessage = @()
#In my testing you should not do more than 4 threads due to conflicts. Unfortunately, I had to abandon the parallel.
#$studentsCSV | ForEach-Object -ThrottleLimit $Threads -TimeoutSeconds 0 -Parallel {
$studentsCSV | ForEach-Object {

    #if ($usererror -eq $True) { continue }
    
    #running in parallel requires us to pull in the variables and functions again.
    #. ./settings.ps1
    #. ./z_functions.ps1

    $student = $PSItem
    $studentId = [int]$student.'Student_id'

    #excluded accounts shouldn't have anything done to them.
    if ($excludedStudentIDs -contains $studentId) { return }

    $givenName = $student.'First_name'
    $surName = $student.'Last_name'

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

    $fullName = $firstName + ' ' + $lastName
   
    switch ($student.'Grade') {
        'Prekindergarten' { $grade = 'PK' }
        'Kindergarten' { $grade = 'K' }
        'KG' { $grade = 'K' }
        'KF' { $grade = 'K' }
        default { $grade = $student.'Grade' }
    }

    $gender = $student.'Gender'
    $buildingNumber = [int]$student.'School_id'
    $buildingShortName = $validschools.$([int]$student.'School_id')
    $gradyr = $student.'Student_gradyr'

    if (($null -eq $gradyr) -or ($gradyr -eq '')) {
        Write-Host "Error: Missing GradYR for student $($studentId). Skipping Student as we need that data."
        errorCount++
        errorMessage += @("Error: Missing GradYR for student $($studentId). Skipping Student as we need that data.")
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
            $username = "$($firstName + '.' + $($lastName.Substring(0,3)))"
            $principalName = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            $emailAddress = $username + $([string]$gradyr).Substring(2,2) + $emailDomain
            if ($username.Length -gt 18) { $username = "$($($firstName.Substring(0,14)) + '.' + $($lastName.Substring(0,3)))" } #shortens first name to meet 20 characters
            $username = $username + $([string]$gradyr).Substring(2,2)
        }
    }

    if ($($homeDirectoryRoot.$buildingNumber)) {
        $homeDirRoot = $homeDirectoryRoot.$buildingNumber + $($buildingShortName) + '\'
        if (@(1,2,3) -contains $adstructure) {
            $homeDirPath = "$($homeDirRoot)$($gradyr)"
        } elseif (@(2,3,4) -contains $adStructure) {
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
                    if (@(2,4,6) -contains $stuTemplate) {
                        Write-Host "Staging: $($studentId) This account has a different graduation year than before. This account will be renamed to reflect the graduation year change." -ForegroundColor Green
                    }
                }
            }
            return
        }

        if ($VerboseStudent) { Write-Host "Verbose: Account already exists for $($fullname)." -ForegroundColor Yellow }

        if ($StopAfterXExisting -ge 1) {
            $existingAccountsCount++
            if ($existingAccountsCount -gt $StopAfterXExisting) {
                Write-Host "Info: Threshold of $StopAfterXExisting on processing existing accounts has been met. Breaking from modifying students."
                break
            }
        }

        #Enable Account if Disabled (returning student)
	    if ($existingAccount.Enabled -eq $False) {
            try {
                Write-Host "Info: Enabling account $($existingAccount.'DistinguishedName')" -ForeGroundColor Yellow
                Set-AdAccountPassword -Identity $existingAccountGUID -Reset -NewPassword (ConvertTo-SecureString "$password" -AsPlainText -Force)
                Set-AdUser -Identity $existingAccountGUID -ChangePasswordAtLogon $False
                Set-ADAccountControl -Identity $existingAccountGUID -CannotChangePassword $False
                Enable-ADAccount -Identity $existingAccountGUID

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
        if ($stuGradYrRename) {
            #if ($existingAccount.Fax -ne $gradyr) { #If a student is returning and their gradyr wasn't in the .Fax field it caused a failure. So we need to actually check the username.
            if ( ($($existingAccount.UserPrincipalName).Substring($($($existingAccount.UserPrincipalName).Split('@')[0]).Length - 2,2)) -ne $([string]$gradyr).Substring(2,2) ) {
                if (@(2,4,6,7) -contains $stuTemplate) {
                    $existingAccount.surname = '' #Changing the name will force a rename in the next code block which should fix the gradyr change.
                }
            }
        }

        #Rename Account if name mismatch. Only on Firstname and Given name as username may change on duplicate.
        if (($existingAccount.GivenName -ne "$givenName")`
         -or ($existingAccount.surname -ne "$surName")`
         #-or ($existingAccount.displayname -ne "$fullname")`
         #-or ($existingAccount.name -ne "$principalName")`
         #-or ($existingAccount.SamAccountName -ne "$username")`
         #-or ($existingAccount.Mail -ne "$emailAddress")`
        ) {
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
                Write-Host "Error: Failed to rename $oldName to $username"
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

        #Check if account is in the restricted OU.
        if ($existingOU -like "*OU=Restricted*") {
            #if current Restricted OU is not like the calculated OU then we have to assume a building change.
            if ($ou -notlike "*$(($existingAccount.DistinguishedName).split(',')[(($($existingAccount.DistinguishedName).Split(',').indexof('OU=Restricted'))+1)..($($existingAccount.DistinguishedName).Split(',').Length)] -join ',')") {
                #Set $ou to the new building but keep it in the restricted OU for the new building.
                $ou = "OU=Restricted,$($($ou).split(',')[1..($($ou).Length -1 )] -join ',')"
            } else {
                $ou = $existingOU
                Write-Host "Info: Not moving student $($existingAccount.DistinguishedName) from Restricted OU."
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

    } else {
        #we need to create a new account for the student. If we have a conflict we need to decide how to cope with it.
        #if ($staging) { Write-Host "Staging: Create an account for $username, $fullname, $gradyr, $buildingShortName, at $ou" -ForeGroundColor Yellow }

        if ($SkipNewStudents) { return }
        if ($VerboseStudent) { Write-Host "Info: Need to verify we can create a new account for $username, $fullname, $gradyr, $buildingShortName, at $ou" -ForeGroundColor Yellow }

        if ($StopAfterXNew -ge 1) {
            $newAccountsCount++
            if ($newAccountsCount -gt $StopAfterXNew) {
                Write-Host "Info: Threshold of $StopAfterXNew on processing new accounts has been met. Breaking from modifying students."
                break
            }
        }

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
                    & $currentPath\..\gam\gam.exe create user $emailAddress
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

                    #testing
                    # $newStudents | ForEach-Object {
                    #     $filename = "$($currentPath)\passwords\$($PSItem.'Building')-passwords.csv"
                    #     "$($PSItem.'Student_id'),$($PSItem.'Name'),$($PSItem.'emailAddress'),$($PSItem.'password')`r`n" | Out-File $filename -Append -Force -NoNewLine -Encoding UTF8
                    # }
                    # break

            }
            
            $homeDirRoot = $null; $homeDirPath = $null; $homeDir = $null
        } catch {
            Write-Host "Error: Failed to create account for $($fullName), $($username), $($password), at $($ou). Try running script again in case this is a duplicate user error." -ForegroundColor Red
            Write-Host $Error[0].Exception.GetType().fullname
            $homeDirRoot = $null; $homeDirPath = $null; $homeDir = $null
            $errorCount++
            $errorMessage += @("Error: Failed to create account for $($fullName), $($username), at $($ou). Try running script again in case this is a duplicate user error.")
            return
        }
        
    }

}
}

#####################################################################
# Write Passwords to CSV and email new ones
#####################################################################

#if ResetAllPasswords was set dump to CSV
if ($ResetAllPasswords) {
    $passwordhashtable.Keys | ForEach-Object {
        Out-File -InputObject $passwordhashtable.$PSItem -FilePath "$($currentPath)\passwords\$($PSItem)-passwords.csv" -Append -Force -NoNewLine -Encoding UTF8
    }
}

#Write new student passwords to CSV. [note to self: convert this to the pwsh7 and export as csv directly.]
$newStudents | ForEach-Object {
    $filename = "$($currentPath)\passwords\$($PSItem.'Building')-passwords.csv"
    "$($PSItem.'Student_id'),$($PSItem.'Name'),$($PSItem.'emailAddress'),$($PSItem.'password')`r`n" | Out-File $filename -Append -Force -NoNewLine -Encoding UTF8
}

#####################################################################
# Student Grade and Building Distribution Groups
#####################################################################

#Domain wide student group. This is a security group for shares, gpo filtering, etc. not a distribution group. You can manually add the email address if you want it to be a distribution group.
$groupUsers = @()
switch ($adStructure) {
    1 {
        $ou = "ou=Students,$($domain)"
        $groupUsers += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=Students,$($domain)" | Select-Object -ExpandProperty SamAccountName
    }
    2 {
        $ou = "ou=Students,$($domain)"
        $validschools.values | ForEach-Object {
            if ($Staging) { return }
            $groupUsers += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=$($PSItem),ou=Students,$($domain)" | Select-Object -ExpandProperty SamAccountName
        }
    }
    3 {
        $ou = "cn=Users,$($domain)" #where else could we put this?
        $validschools.values | ForEach-Object {
            if ($Staging) { return }
            $groupUsers += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=Students,ou=$($PSItem),$($domain)" | Select-Object -ExpandProperty SamAccountName
        }
    }
    4 {
        $ou = "ou=Students,$($domain)"
        $groupUsers += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=Students,$($domain)" | Select-Object -ExpandProperty SamAccountName
    }
    5 {
        $ou = "ou=Students,$($domain)"
        $validschools.values | ForEach-Object {
            if ($Staging) { return }
            $groupUsers += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=$($PSItem),ou=Students,$($domain)" | Select-Object -ExpandProperty SamAccountName
        }
    }
    6 {
        $ou = "cn=Users,$($domain)" #where else could we put this?
        $validschools.values | ForEach-Object {
            if ($Staging) { return }
            $groupUsers += Get-ADUser -Filter "(EmployeeNumber -like ""$($studentIDPrefix)*"") -and (Enabled -eq 'True')" -SearchBase "ou=Students,ou=$($PSItem),$($domain)" | Select-Object -ExpandProperty SamAccountName
        }
    }
}

try {
    $ADGroup = Get-ADGroup "students" -ErrorAction SilentlyContinue
    if (-Not($staging)) {
        if ($studentEmailGroupOwners.students) {
            $groupOwners = $studentEmailGroupOwners.students
            Set-ADGroup -Identity $ADGroup -Replace @{"wbemPath" = $groupOwners}
        } else {
            Write-Host "Info: No group owners defined for group $groupName. Please check the" '$studentEmailGroupOwners' "variable in settings.ps1. Otherwise manually make updates in AD to wbemPath attribute."
        }
        Set-ADGroupMembershipOnly "students" $groupUsers $excludedAccounts
    }
} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    #We can create the group here
    if ($staging) {
        Write-Host "Staging: Create new group "students" at $ou" -ForeGroundColor Yellow
        New-ADGroup -Name "students" -Path $ou -GroupScope Global -WhatIf
    } else {
        New-ADGroup -Name "students" -Path $ou -GroupScope Global -OtherAttributes @{'wbemPath'="noreply$($teacherDomain)"}
        if ($studentEmailGroupOwners.students) {
            $groupOwners = $studentEmailGroupOwners.students
            Get-ADGroup -Filter { name -eq "students" } | Set-ADGroup -Replace @{"wbemPath" = $groupOwners}
        } else {
            Write-Host "Info: No group owners defined for group $groupName. Please check the" '$studentEmailGroupOwners' "variable in settings.ps1. Otherwise manually make updates in AD to wbemPath attribute."
        }
        Set-ADGroupMembershipOnly "students" $groupUsers $excludedAccounts
    }
} catch {
    Write-Host "Error: Can not find the group ""students"" and did not get a Not Found Exception. Error is $($PSItem.Exception.Message)" -ForegroundColor Red
    Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Can not find ""students"" and did not get a Not Found Exception. Error is $($PSItem.Exception.Message)"
    exit(1)
}

#grade level distribution lists.
$studentsCSV | Group-Object -Property 'School_id'| ForEach-Object { $PSItem.Group | Select-Object -Property School_id,Grade -Unique } | ForEach-Object {

    $schoolGrade = $PSItem
    $buildingShortName = $validschools.([int]$schoolGrade.'School_id')

    switch ($adStructure) {
        1 { $ou = "ou=Students,$($domain)" }
        2 { $ou = "ou=Students,$($domain)" }
        3 { $ou = "ou=$($buildingShortName),$($domain)" }
        4 { $ou = "ou=Students,$($domain)" }
        5 { $ou = "ou=Students,$($domain)" }
        6 { $ou = "ou=$($buildingShortName),$($domain)" }
    }

    $grade = $schoolGrade.'Grade'

    switch ($grade) {
        'Prekindergarten' { $grade = 'PK' }
        'Kindergarten' { $grade = 'K' }
        'KG' { $grade = 'K' }
        'K' { $grade = 'K' }
        'KF' { $grade = 'K' }
        'PK' { $grade = 'PK' }
        1 { $grade = [string]$([int]$grade) }
        2 { $grade = [string]$([int]$grade) }
        3 { $grade = [string]$([int]$grade) }
        4 { $grade = [string]$([int]$grade) }
        5 { $grade = [string]$([int]$grade) }
        6 { $grade = [string]$([int]$grade) }
        7 { $grade = [string]$([int]$grade) }
        8 { $grade = [string]$([int]$grade) }
        9 { $grade = [string]$([int]$grade) }
        10 { $grade = [string]$([int]$grade) }
        11 { $grade = [string]$([int]$grade) }
        12 { $grade = [string]$([int]$grade) }
        default { break } #can't find a match so break
    }

    if ($stuEmailDomain.$([int]$student.'School_id')) {
        $emailDomain = $stuEmailDomain.([int]$schoolGrade.'School_id')
    } else {
        $emailDomain = $stuEmailDomain.'default'
    }

    $groupName = "students-$($buildingShortName)-$($grade)"
    $groupMail = $groupName + $emailDomain

    #Create distribution lists for each grade level at the buildings. Example: students-ghs-12
    $groupUsers = Get-ADUser -Filter "homePhone -eq ""$([string]$grade)"" -and Office -eq ""$buildingShortName"" -and Enabled -eq 'True'" | Select-Object -ExpandProperty SamAccountName
    try {
        $ADGroup = Get-ADGroup $groupName -ErrorAction SilentlyContinue
        if (-Not($staging)) {
            if ($studentEmailGroupOwners.$buildingShortName) {
                $groupOwners = $studentEmailGroupOwners.$buildingShortName
                if ($studentEmailGroupOwners.$grade) { $groupOwners += $studentEmailGroupOwners.$grade }
                if ($studentEmailGroupOwners."$($buildingShortName)-$($grade)") { $groupOwners += $studentEmailGroupOwners."$($buildingShortName)-$($grade)" }
                Set-ADGroup -Identity $ADGroup -Replace @{"wbemPath" = $groupOwners}
            } else {
                Write-Host "Info: No group owners defined for group $groupName. Please check the" '$studentEmailGroupOwners' "variable in settings.ps1. Otherwise manually make updates in AD to wbemPath attribute."
            }
            Set-ADGroupMembershipOnly $groupName $groupUsers $excludedAccounts
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        #We can create the group here
        if ($staging) {
            Write-Host "Staging: Create new group $groupName at $ou" -ForeGroundColor Yellow
            New-ADGroup -Name $groupName -Path $ou -GroupScope Global -OtherAttributes @{'mail'=$groupMail} -WhatIf
        } else {
            New-ADGroup -Name $groupName -Path $ou -GroupScope Global -OtherAttributes @{'mail'=$groupMail;'wbemPath'="noreply$($teacherDomain)"}
            if ($studentEmailGroupOwners.$buildingShortName) {
                $groupOwners = $studentEmailGroupOwners.$buildingShortName
                if ($studentEmailGroupOwners.$grade) { $groupOwners += $studentEmailGroupOwners.$grade }
                if ($studentEmailGroupOwners."$($buildingShortName)-$($grade)") { $groupOwners += $studentEmailGroupOwners."$($buildingShortName)-$($grade)" }
                Get-ADGroup -Filter { name -eq $groupName } | Set-ADGroup -Replace @{"wbemPath" = $groupOwners}
            } else {
                Write-Host "Info: No group owners defined for group $groupName. Please check the" '$studentEmailGroupOwners' "variable in settings.ps1. Otherwise manually make updates in AD to wbemPath attribute."
            }
            Set-ADGroupMembershipOnly $groupName $groupUsers $excludedAccounts
        }
    } catch {
        Write-Host "Error: Can not find $($groupname) and did not get a Not Found Exception. Error is $($PSItem.Exception.Message)" -ForegroundColor Red
        Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Can not find $($groupname) and did not get a Not Found Exception. Error is $($PSItem.Exception.Message)"
        exit(1)
    }
}

#building level distribution list
$validschools.keys | ForEach-Object {

    $buildingShortName = $validschools.([int]$PSItem)
 
    switch ($adStructure) {
        1 { $ou = "ou=Students,$($domain)" }
        2 { $ou = "ou=Students,$($domain)" }
        3 { $ou = "ou=$($buildingShortName),$($domain)" }
        4 { $ou = "ou=Students,$($domain)" }
        5 { $ou = "ou=Students,$($domain)" }
        6 { $ou = "ou=$($buildingShortName),$($domain)" }
    }

    if ($stuEmailDomain.$([int]$PSItem)) {
        $emailDomain = $stuEmailDomain.([int]$PSItem)
    } else {
        $emailDomain = $stuEmailDomain.'default'
    }

    $groupName = "students-$($buildingShortName)"
    $groupMail = $groupName + $emailDomain

    $groupUsers = Get-ADUser -Filter "Office -eq ""$buildingShortName"" -and Enabled -eq 'True'" | Select-Object -ExpandProperty SamAccountName
    try {
        $ADGroup = Get-ADGroup $groupName -ErrorAction SilentlyContinue
        if (-Not($staging)) {
            if ($studentEmailGroupOwners.$buildingShortName) {
                $groupOwners = $studentEmailGroupOwners.$buildingShortName
                Set-ADGroup -Identity $ADGroup -Replace @{"wbemPath" = $groupOwners}
            } else {
                Write-Host "Info: No group owners defined for group $groupName. Please check the" '$studentEmailGroupOwners' "variable in settings.ps1. Otherwise manually make updates in AD to wbemPath attribute."
            }
            Set-ADGroupMembershipOnly $groupName $groupUsers $excludedAccounts
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        #We can create the group here
        if ($staging) {
            Write-Host "Staging: Create new group $groupName at $ou" -ForeGroundColor Yellow;
            New-ADGroup -Name $groupName -Path $ou -GroupScope Global -OtherAttributes @{'mail'=$groupMail} -WhatIf
        } else {
            New-ADGroup -Name $groupName -Path $ou -GroupScope Global -OtherAttributes @{'mail'=$groupMail;'wbemPath'="noreply$($teacherDomain)"}
            if ($studentEmailGroupOwners.$buildingShortName) {
                $groupOwners = $studentEmailGroupOwners.$buildingShortName
                Get-ADGroup -Filter { name -eq $groupName } | Set-ADGroup -Replace @{"wbemPath" = $groupOwners}
            } else {
                Write-Host "Info: No group owners defined for group $groupName. Please check the" '$studentEmailGroupOwners' "variable in settings.ps1. Otherwise manually make updates in AD to wbemPath attribute."
            }
            Set-ADGroupMembershipOnly $groupName $groupUsers $excludedAccounts
        }
    } catch {
        Write-Host "Error: Can not find $($groupname) and did not get a Not Found Exception. Error is $($PSItem.Exception.Message)" -ForegroundColor Red
        Send-EmailNotification -subject "Automated_Students: Failure" -body "Error: Can not find $($groupname) and did not get a Not Found Exception. Error is $($PSItem.Exception.Message)"
        exit(1)
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
        if (Test-Path $currentPath\x_PostProcessingAutomatedStudents.ps1) {
            . $currentPath\x_PostProcessingAutomatedStudents.ps1
        }
    } else {
        Write-Host "Info: Staging specified. Will not run x_PostProcessingAutomatedStudents.ps1."
    }
} else {
    Write-Host "Info: DisablePostProcessingScript specified. Will not run x_PostProcessingAutomatedStudents.ps1."
}

Stop-Transcript
