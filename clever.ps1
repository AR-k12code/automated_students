#Requires -Version 7.0

Param(
    [Parameter(Mandatory=$false)][switch]$SkipUploadingFiles #Skip Uploading
)

###########################################################
# Clever Upload Script from PSSQLite
# Craig Millsap, Gentry Public Schools, cmillsap@gentrypioneers.com, 9/2020
###########################################################

# This script requires the CognosDownload.ps1 script by Brian Johnson
# pscp.exe and puttygen.exe, private keys for clever and scriptk12.
# You must generate a ssh key for Clever and upload the public key at https://schools.clever.com/sync/settings

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

Import-Module PSSQLite
. $currentPath\settings.ps1

if (-Not(Test-Path $currentPath\clever)) { New-Item -ItemType Directory clever }
if (-Not(Test-Path $currentPath\clever\customcontacts)) { New-Item -ItemType Directory clever\customcontacts }
if (-Not(Test-Path $currentPath\bin)) { New-Item -ItemType Directory bin }
if (-Not(Test-Path $currentPath\bin\pscp.exe) -or -Not(Test-Path $currentPath\bin\puttygen.exe)) {
    Write-Host "You need the pscp.exe from https://the.earth.li/~sgtatham/putty/latest/w64/pscp.exe in the bin folder."
    Write-Host "You need the puttygen.exe from https://the.earth.li/~sgtatham/putty/latest/w64/puttygen.exe in the bin folder."
    try {
        Invoke-WebRequest -Uri 'https://the.earth.li/~sgtatham/putty/0.70//w64/pscp.exe' -OutFile "$currentPath\bin\pscp.exe"
        Invoke-WebRequest -Uri 'https://the.earth.li/~sgtatham/putty/latest/w64/puttygen.exe' -OutFile "$currentPath\bin\puttygen.exe"
    } catch {
        exit(1)
    }
}
if (-Not(Test-Path $currentPath\keys)) { New-Item -ItemType Directory keys }
if (-Not(Test-Path $currentPath\keys\clever.ppk)) {
    Write-Host "Clever private key not found! Use Putty Key Generator to create a new key and upload at https://schools.clever.com/sync/settings" -ForegroundColor RED
    Start-Process -FilePath "$currentPath\bin\puttygen.exe"
    exit(1)
}
if ($eschoolusername -notmatch [regex]('\d{4}.*')) {
    Write-Host "Invalid eSchool username. Must be LEA#username."
    exit(1)
}
if ($cleverusername.length -eq 0) {
    Write-Host "Invalid clever login username. Get your username from https://schools.clever.com/sync/settings"
    exit(1)
}
if ($cleverserver.length -eq 0) {
    Write-Host "Invalid clever login username. Get your username from https://schools.clever.com/sync/settings"
    exit(1)
}

###########################################################
# Pull Term Data from imported Sections and select currentTerm.
###########################################################
$term1 = Invoke-SqliteQuery -DataSource $database -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 1 LIMIT 1"
$term2 = Invoke-SqliteQuery -DataSource $database -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 2 LIMIT 1"
$term3 = Invoke-SqliteQuery -DataSource $database -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 3 LIMIT 1"
$term4 = Invoke-SqliteQuery -DataSource $database -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 4 LIMIT 1"

$today = Get-Date
if (([Int]$(Get-Date -Format MM) -gt 7) -And ($today -le [datetime]$term1.Term_end)) {
    $currentTerm = 1
} elseif (($today -gt [datetime]$term1.Term_end) -And ($today -le [datetime]$term2.Term_end)) {
    $currentTerm = 2
} elseif (($today -gt [datetime]$term2.Term_end) -And ($today -le [datetime]$term3.Term_end)) {
    $currentTerm = 3
} else {
    $currentTerm = 4 #why do math when its the only option left?
}

#Pull data from PSSQLite and create CSVs for Clever Upload.
Write-Host "Info: Querying data from SQL and building CSVs for Clever..." -ForegroundColor YELLOW
try {
    $students = Invoke-SqliteQuery -Database $database -Query "SELECT * FROM students" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp
    
    #I completely disagree with this action. School email addresses should already be pushed back into eSchool and we should be getting correct data from eSchool.
    #But No, we have this hack job. This is the most passive aggressive away I can express my anger about it. In my comments. Gah, no relief!
    $students | ForEach-Object {
        $email = $PSItem.Student_email
        #If its blank or doesn't match a student email domain specified in settings.ps1 then we go looking in AD.
        if (($null -eq $email) -or ($email -eq '') -or ($stuEmailDomain.Values -notcontains "@$($email.Split('@')[1])")) {
            #This shouldn't fail because we should have already generated student AD accounts.
            Write-Host "Notify:",$PSItem.Student_id,"is either missing or has an invalid email address on their eSchool record. Pulling from AD." -ForegroundColor Yellow
            $PSItem.Student_email = Get-ADUser -Filter "EmployeeNumber -eq $($PSItem.Student_id)" -Properties mail | Select-Object -ExpandProperty mail
        }
    }
    
    $students | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\students.csv -Force

    $sections = Invoke-SqliteQuery -Database $database -Query "SELECT * FROM sections WHERE Term_name = $currentTerm" -ErrorAction 'STOP'
    $sections += Invoke-SqliteQuery -Database $database -Query "SELECT * FROM sections WHERE Term_name != $currentTerm" -ErrorAction 'STOP'
    $sections | Select-Object -Property * -ExcludeProperty Timestamp | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\sections.csv -Force

    $enrollments = Invoke-SqliteQuery -Database $database -Query "SELECT School_id,Section_id,Student_id FROM schedules" -ErrorAction 'STOP'
    $enrollments | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\enrollments.csv -Force

    $teachers = Invoke-SqliteQuery -Database $database -Query "SELECT * FROM teachers" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp
    $teachers | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\teachers.csv -Force

    $schools = Invoke-SqliteQuery -Database $database -Query "SELECT * FROM schools" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp
    $schools | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\schools.csv -Force

    #Contact_id,Student_id,Contact_firstname,Contact_lastname,Contact_type,Contact_phonenumber,Contact_relationship,Contact_phonetype,Contact_email
    #contact_id,student_id,first_name,last_name,type,phone,relationship,phone_type,email
    $contacts = Invoke-SqliteQuery -Database $database -Query "SELECT * FROM contacts WHERE Contact_phonetype IN ('C','C1','C2')" -ErrorAction 'STOP' | Select-Object -Property `
        @{Name='Contact_id';Expression={$PSItem.'Contact_id'}},
        @{Name='Student_id';Expression={$PSItem.'Student_id'}},
        @{Name='first_name';Expression={$PSItem.'Contact_firstname'}},
        @{Name='last_name';Expression={$PSItem.'Contact_lastname'}},
        @{Name='type';Expression={$PSItem.'Contact_type'}},
        @{Name='phone';Expression={$PSItem.'Contact_phonenumber'}},
        @{Name='relationship';Expression={$PSItem.'Contact_relationship'}},
        @{Name='phone_type';Expression={$PSItem.'Contact_phonetype'}},
        @{Name='email';Expression={$PSItem.'Contact_email'}}
    $contacts | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\customcontacts\contacts.csv -Force
} catch {
    write-host "ERROR: Failed to query SQL and create CSVs." -ForegroundColor RED
    exit(1)
}

#Do we need to modify these files further before sending them to Clever?
if (Test-Path $currentPath\x_InterimProcessingClever.ps1) {
    . $currentPath\x_InterimProcessingClever.ps1
}

#SCP upload the files
try {
    if (-Not($SkipUploadingFiles)) {
        Write-Host "Info: Uploading files to Clever..." -ForegroundColor YELLOW
        Start-Process -FilePath "$currentPath\bin\pscp.exe" -ArgumentList "-r -i $currentpath\keys\clever.ppk -r $currentpath\clever\ $($cleverusername)@$($cleverserver):" -PassThru -Wait -NoNewWindow
    } else {
        Write-Host "Info: You have requested to skip uploading files."
    }
} catch {
    write-Host "ERROR: Failed to properly upload files to clever." -ForegroundColor RED
    exit(1)
}

exit