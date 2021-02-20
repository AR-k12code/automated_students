#Requires -Version 7.0

Param(
    [Parameter(Mandatory=$false)][switch]$ClearDatabase, #Drops all tables so it can be recreated.
    [Parameter(Mandatory=$false)][switch]$SkipDownloadingReports, #Do not download updated report files using the CognosDownloader
    [Parameter(Mandatory=$false)][switch]$DownloadFilesOnly, #Download new files from Cognos then exit.
    [Parameter(Mandatory=$false)][switch]$SkipTableCleanup, #Do not remove rows that doen't match current timestamp.
    [Parameter(Mandatory=$false)][switch]$DisablePostProcessingScript, #Do not run the final script.
    [Parameter(Mandatory=$false)][switch]$SkipRemoteCheck #If remote check is enabled but you want to run anyways.
)

###############################################################################
# Student SQL Database to pull local queries from.
# Craig Millsap, Gentry Public Schools, cmillsap@gentrypioneers.com, 9/2020
# SQLite Database Browser: http://sqlitebrowser.org
# This script requires modified Clever reports and a few additional custom
# queries from Cognos to work. I will do my best to keep the reports/queries
# in a shared place in Cognos.
# Do not make fun of my tables. I know they don't meet nf standards. Live with it.
###############################################################################
$scriptVersion = 1.3
$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

if (-Not(Test-Path $currentPath\logs)) { New-Item -ItemType Directory -Path $currentPath\logs }
$logfile = "$currentPath\logs\automated_sqlite_$(get-date -f yyyy-MM-dd-HH-mm-ss).log"
try {
    Start-Transcript($logfile)
} catch {
    Stop-TranScript; Start-Transcript($logfile)
}

if ($PSVersionTable.PSVersion.Major -lt 7) {
  write-host "This script requires Powershell 7 or higher." -foregroundcolor RED
  exit(1)
}

#PSSQLite Module is required.
if (Get-Module -ListAvailable | Where-Object {$PSItem.name -eq "PSSQLite"}) {
  Try { Import-Module PSSQLite } catch { Write-Host "Error: Unable to load module PSSQLite." -ForegroundColor RED; exit(1) }
} else {
  Write-Host 'PSSQLite Module not found!'
  if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically install it? y/n')) {
      try {
          Install-Module -Name PSSQLite -Scope AllUsers -Force
          Import-Module PSSQLite
      } catch {
          write-host 'Failed to install PSSQLite Module.' -ForegroundColor RED
          exit(1)
      }
  } else {
      exit(1)
  }
}

#Pull in Variables
if (Test-Path $currentPath\settings.ps1) {
  . $currentPath\settings.ps1
} else {
  write-host "Error: Missing settings.ps1 file. Please read documentation." -ForegroundColor Red; exit(1)
}

#Import Needed Functions
if (Test-Path $currentPath\z_functions.ps1) {
  . $currentPath\z_functions.ps1
} else {
  write-host "Error: Missing z_functions.ps1 file. Please read documentation." -ForegroundColor Red; exit(1)
}

#Remote Stop & Update
if ($remoteCheck -and (-Not($SkipRemoteCheck))) {
  try {
    Write-Host "Info: Checking remote server to ensure it is safe to proceed."
    $remoteResponse = Invoke-RestMethod -Uri $remoteCheckURL
    if ($remoteResponse.status -ne "OK") {
      Write-Host "Warning: Remote Server has indicated that it is not safe to proceed." -ForegroundColor Yellow
      exit(1)
    }
    if ($scriptVersion -lt $remoteResponse.version) {
      Write-Host "Info: There is a new version of these scripts available. Please run 'git pull' at $currentPath to upgrade."
    }
  } catch {
    Write-Host "Error: You have specified that we should check the remote server to ensure its safe to proceed. Exiting since we are unable to verify."
    exit(1)
  }
}

$timestamp = [int64](Get-Date -UFormat %s)
if ($RetainEnrollmentsDays -ge 1) {
  $timestampRetain = [long] (Get-Date -Date (((Get-Date).AddDays(-$RetainEnrollmentsDays)).ToUniversalTime()) -UFormat %s)
}
$errorMessage = @()

#cleanup old log files.
Get-ChildItem $currentPath\logs | Where-Object {$PSItem.LastWriteTime -lt (Get-Date).AddDays(-$keeplogfiledays)} | Remove-Item -Force

#####################################################################
# Required folder for Cognos Reports then loop through and download
# required reports. We are not checking here for failures.
# The CognosDownload.ps1 script does error checking and will send an email notification.
# If downloading updated reports fails the script will continue with existing files.
#####################################################################
if (-Not(Test-Path $currentPath\files)) { New-Item -ItemType Directory -Path $currentPath\files }

if ([int](Get-Date -Format MM) -ge 6) {
    $schoolyear = [int](Get-Date -Format yyyy) + 1
} else {
    $schoolyear = [int](Get-Date -Format yyyy)
}

$reports = @{
    'enrollments' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'report'}
    'schools' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'report'}
    'sections' = @{ 'arguments' = "p_year=$([string]$schoolyear)"; 'folder' = 'automation'; 'type' = 'report'}
    'students' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'report'}
    'teachers' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'report'}
    'students_extras' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'query'}
    'contacts' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'query'}
    'facultyids' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'query'}
    'activities' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'query'}
    'transportation' = @{ 'arguments' = ''; 'folder' = 'automation'; 'type' = 'query'}
}

if ($SkipDownloadingReports -eq $False) {

  if (-Not(Test-Path "$currentPath\..\CognosDownload.ps1")) {
    Write-Host "ERROR: CognosDownload script is missing." -ForegroundColor RED
    if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically download it? y/n')) {
        try {
            #Waiting on Pull Request to address PWSH7 incompatibility.
            #Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/AR-k12code/CognosDownloader/master/CognosDownload.ps1' -OutFile "$currentPath\..\CognosDownload.ps1"
            Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/carbm1/CognosDownloader/master/CognosDownload.ps1' -OutFile "$currentPath\..\CognosDownload.ps1"
        } catch {
            write-host "ERROR: Failed to download CognosDownloader script." -ForegroundColor RED
            exit(1)
        }
    } else {
        Write-Host "ERROR: The CognosDownloader is required for this script unless you specify -SkipDownloadingReports and manually download your CSV files to $currentPath\files" -ForegroundColor RED
        exit(1)
    }
  }

    $reports.keys | ForEach-Object {

        $report = $reports.$PSItem

        $command = "$($currentPath)\..\CognosDownload.ps1" -replace ' ','` ' #escape spaces in the filepath.
        $command += " -report $($PSItem) -RunReport -savepath ""$currentPath\files\"" -username $eSchoolUsername -espdsn $eSchooldsn -extension csv -reportwait $reportWait"
        
        #params
        if ($($report.'arguments').length -ge 1) { $command += " -reportparams $($report.'arguments')" }
        if ($($report.'folder') -ge 1) { $command += " -cognosfolder $($report.'folder')" }
        if ($($report.'type') -eq "report") { $command += " -ReportStudio" }
        
        write-host "INFO: Invoking Cognos Downloader script to pull $PSItem"
        Invoke-Expression $command #write-host $command
        
        if ($LASTEXITCODE -ge 1) {
          Write-Host "Error: Failed to properly download Cognos report $PSItem"
          $errorCount++
          $errorMessage += @("Error: Failed to properly download Cognos report $PSItem")
        }
    }
}

if ($errorCount -ge 1) {
  $errorMessageString = $errorMessage -join "`r`n"
  if ($sendMailNotifications) {
    Send-EmailNotification -subject "Automated_Students: Cognos Report Failure" -body "There were $errorCount errors detected while processing student accounts.`r`n$($errorMessageString)`r`nPlease inspect the log file $($logfile) for more details.`r`n"
  }
}


#####################################################################
# Verify all required CSV files exist then import and merge them.
#####################################################################
write-host "Info: Verifying that all CSV files exist."
$missingFileNames = @()
$reports.Keys | ForEach-Object {
    $file = "$($PSItem).csv"
    if (-Not(Test-Path ".\files\$($file)")) {
        $missingFileNames += $file
    }
}
if ($missingFileNames.Length -ge 1) {
    $missingFiles = ($($missingFileNames) -join (','))
    Write-Host "ERROR: You are missing the following CSV files: $missingFiles" -ForegroundColor RED  
    exit(1)
}

# Sanitize files by removing any special characters
write-host "Info: Sanitizing and removing special characters from import files."
$reports.Keys | ForEach-Object {
    $fileContents = Get-Content -Encoding utf8 ".\files\$($PSItem).csv" | Remove-StringLatinCharacter
    $fileContents | Set-Content -Encoding utf8 ".\files\$($PSItem).csv"
    $fileContents = $NULL
}

#Verify that the files imported have at least a School_id or Student_id property and a count greater than 3. I don't know any schools with less than 3 campuses. #activites.csv is an exception.
write-host "Info: Verifying each file has a School_id or Student_id column."
$reports.Keys | ForEach-Object {
  if ($PSItem -ne 'activities') {
    try {
      $verifyContents = Import-CSV ".\files\$($PSItem).csv"
      if ($verifyContents -eq $NULL) {
        write-host "ERROR: $($PSItem) file must be empty." -ForegroundColor RED
        exit(1)
      }
      $headers = $verifyContents | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
      if (($headers -contains 'School_id') -OR ($headers -contains 'Student_id')) {
        #Need to verify that there is actual data in the files.
        if ($($verifyContents | Measure-Object).Count -lt 3) {
          write-host "ERROR: No data found in $($PSItem)." -ForegroundColor RED
          exit(1)
        }
      } else {
        #File doesn't have the right headers.
        write-host "ERROR: Problem with headers in $($PSItem)." -ForegroundColor RED
        exit(1)
      }
    } catch {
      write-host "ERROR: Problem importing $($PSItem) and verifying data." -ForegroundColor RED
      exit(1)
    }
  }
}

if ($DownloadFilesOnly) { exit }

#Quick way to clear out database.
if ($ClearDatabase) {
  write-host "Info: Clear database was specified. Dropping all tables except for disabled_students and enrollments_archived." -Foregroundcolor YELLOW
  Invoke-SqliteQuery -DataSource $database -Query 'drop table if exists schools;
    drop table if exists teachers;
    drop table if exists students;
    drop table if exists sections;
    drop table if exists enrollments;
    drop table if exists sections_grouped;
    drop table if exists enrollments_grouped;
    drop table if exists schedules;
    drop table if exists passwords;
    drop table if exists students_extras;
    drop table if exists contacts;
    drop table if exists activities;
    drop table if exists transportation;'
  exit
}

###########################################################
#Verify tables exist and create if not.
###########################################################
write-host "Info: Check if database tables exist and create if not."
#Schools Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM schools limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "schools" (
    "School_id"	INTEGER PRIMARY KEY,
    "School_name"	TEXT,
    "School_number"	TEXT,
    "State_id" TEXT,
    "Low_grade"	TEXT,
    "High_grade"	INTEGER,
    "Principal"	TEXT,
    "Principal_email"	TEXT,
    "School_address"	TEXT,
    "School_city"	TEXT,
    "School_state"	TEXT,
    "School_zip"	INTEGER,
    "School_phone"	INTEGER
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Teachers Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM teachers limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "teachers" (
    "School_id"	INTEGER,
    "Teacher_id" TEXT,
    "Teacher_number"	TEXT,
    "State teacher id"	TEXT,
    "Teacher_email"	TEXT collate nocase,
    "First_name"	TEXT,
    "Middle_name"	TEXT,
    "Last_name"	TEXT,
    "Title"	TEXT,
    "Username"	TEXT,
    "Password"	TEXT,
    "Timestamp" INTEGER,
    UNIQUE(School_id,Teacher_id)
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Students Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM students limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "students" (
    "School_id"	INTEGER,
    "Student_id"	INTEGER PRIMARY KEY,
    "Student_number"	INTEGER,
    "State_id"	INTEGER,
    "Last_name"	TEXT,
    "Middle_name"	TEXT,
    "First_name"	TEXT,
    "Grade"	TEXT,
    "Gender"	TEXT,
    "DOB"	TEXT,
    "Race"	TEXT,
    "Hispanic_Latino"	TEXT,
    "Ell_status"	TEXT,
    "Frl_status"	TEXT,
    "Iep_status"	TEXT,
    "Student_street"	TEXT,
    "Student_city"	TEXT,
    "Student_state"	TEXT,
    "Student_zip"	INTEGER,
    "Student_email"	TEXT collate nocase,
    "Contact_relationship"	TEXT,
    "Contact_type"	TEXT,
    "Contact_name"	TEXT,
    "Contact_phone"	TEXT,
    "Contact_email"	TEXT,
    "Username"	TEXT,
    "Password"	TEXT,
    "Home_language" TEXT,
    "Timestamp" INTEGER
    );'
    Invoke-SqliteQuery -DataSource $database -Query $q
}

#Students_disabled Table - CREATE TABLE AS not implimented in PSSQLite module.
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM students_disabled limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "students_disabled" (
    "School_id"	INTEGER,
    "Student_id"	INTEGER PRIMARY KEY,
    "Student_number"	INTEGER,
    "State_id"	INTEGER,
    "Last_name"	TEXT,
    "Middle_name"	TEXT,
    "First_name"	TEXT,
    "Grade"	TEXT,
    "Gender"	TEXT,
    "DOB"	TEXT,
    "Race"	TEXT,
    "Hispanic_Latino"	TEXT,
    "Ell_status"	TEXT,
    "Frl_status"	TEXT,
    "Iep_status"	TEXT,
    "Student_street"	TEXT,
    "Student_city"	TEXT,
    "Student_state"	TEXT,
    "Student_zip"	INTEGER,
    "Student_email"	TEXT collate nocase,
    "Contact_relationship"	TEXT,
    "Contact_type"	TEXT,
    "Contact_name"	TEXT,
    "Contact_phone"	TEXT,
    "Contact_email"	TEXT,
    "Username"	TEXT,
    "Password"	TEXT,
    "Home_language" TEXT,
    "Timestamp" INTEGER
    );'
    Invoke-SqliteQuery -DataSource $database -Query $q
}

#Sections Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM sections limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "sections" (
    "School_id"	INTEGER,
    "Section_id"	INTEGER,
    "Teacher_id" TEXT,
    "Teacher_2_id"	TEXT,
    "Teacher_3_id"	TEXT,
    "Teacher_4_id"	TEXT,
    "Name"	TEXT,
    "Section_number"	INTEGER,
    "Grade"	INTEGER,
    "Course_name"	TEXT,
    "Course_number"	TEXT,
    "Course_description"	TEXT,
    "Period"	TEXT,
    "Subject"	TEXT,
    "Term_name"	INTEGER,
    "Term_start"	TEXT,
    "Term_end"	TEXT,
    "Timestamp" INTEGER,
    UNIQUE(School_id,Section_id,Term_name)
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Sections_archived Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM sections_archived limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "sections_archived" (
    "School_id"	INTEGER,
    "Section_id"	INTEGER,
    "Teacher_id" TEXT,
    "Teacher_2_id"	TEXT,
    "Teacher_3_id"	TEXT,
    "Teacher_4_id"	TEXT,
    "Name"	TEXT,
    "Section_number"	INTEGER,
    "Grade"	INTEGER,
    "Course_name"	TEXT,
    "Course_number"	TEXT,
    "Course_description"	TEXT,
    "Period"	TEXT,
    "Subject"	TEXT,
    "Term_name"	INTEGER,
    "Term_start"	TEXT,
    "Term_end"	TEXT,
    "Timestamp" INTEGER,
    UNIQUE(School_id,Section_id,Term_name)
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Enrollments Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM enrollments limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "enrollments" (
    "School_id"	INTEGER,
    "Section_id"	INTEGER,
    "Student_id"	INTEGER,
    "Marking_period"	INTEGER,
    "Timestamp" INTEGER,
    UNIQUE(School_id,Section_id,Student_id,Marking_period)
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Enrollments_archived Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM enrollments_archived limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "enrollments_archived" (
    "School_id"	INTEGER,
    "Section_id"	INTEGER,
    "Student_id"	INTEGER,
    "Marking_period"	INTEGER,
    "Timestamp" INTEGER,
    UNIQUE(School_id,Section_id,Student_id,Marking_period)
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Passwords Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM passwords limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "passwords" (
    "Student_id"	INTEGER PRIMARY KEY,
    "Student_password"	TEXT
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#students_extras Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM students_extras limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "students_extras" (
    "Student_id" INTEGER PRIMARY KEY,
    "Student_gradyr" INTEGER,
    "Student_nickname" TEXT,
    "Student_homeroom" TEXT,
    "Student_hrmtid" INTEGER,
    "Student_advisor" INTEGER,
    "Student_houseteam" TEXT,
    "Student_haclogin" TEXT,
    "Student_hacpassword" TEXT,
    "Student_contactid" INTEGER,
    "Student_mealstatus" TEXT,
    Timestamp INTEGER
    );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Contacts Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM contacts limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "contacts" (
    "Contact_id" INTEGER,
    "Student_id" INTEGER,
    "Contact_priority" INTEGER,
    "Contact_firstname"	TEXT,
    "Contact_lastname" TEXT,
    "Contact_type" TEXT,
    "Contact_phonenumber"	INTEGER,
    "Contact_relationship" TEXT,
    "Contact_phonetype"	TEXT,
    "Contact_email"	TEXT,
    Timestamp INTEGER,
    UNIQUE (Contact_id,Student_id,Contact_phonetype)
  );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Activities Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM activities limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "activities" (
    "School_id"	INTEGER,
    "Student_id"	INTEGER,
    "Teacher_id" TEXT,
    "Activity_code"	TEXT,
    "Activity_name"	TEXT,
    Timestamp INTEGER,
    UNIQUE (School_id,Student_id,Teacher_id,Activity_code)
  );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Transportation Table
try {
  $test = Invoke-SqliteQuery -DataSource $database -Query "SELECT * FROM transportation limit 1" -ErrorAction 'STOP'
} catch {
  $q = 'CREATE TABLE "transportation" (
    "Student_id" INTEGER,
    "Student_BusNumFrom" TEXT,
    "Student_TravelTypeFrom" TEXT,
    "Student_BusNumTo" TEXT,
    "Student_TravelTypeTo" TEXT,
    Timestamp INTEGER
  );'
  Invoke-SqliteQuery -DataSource $database -Query $q
}

#Sections_Grouped,Enrollments_Grouped, and Schedules Tables are created/generated at the end from whats imported later.

###########################################################
# Import CSVs into Database Tables
###########################################################

write-host "Info: Importing CSV files into tables."
#I want the originally set passwords in the gentrysms database.
if (Test-Path $currentPath\passwords) {
  write-host "Info: Importing passwords CSVs."
  Get-ChildItem $currentPath\passwords\*.csv | ForEach-Object {
    #need to validate correct headers.
      try {
        $headers = Import-Csv $PSitem.fullname | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
        if (($headers -contains 'Student ID') -and ($headers -contains 'Email Address') -and ($headers -contains 'Password')) {
          $passwords += Import-Csv $PSitem.fullname
        }
      } catch {
        return
      }
  }
  $passwords = $passwords | Where-Object { $NULL -ne $PSItem.'Student ID' } | Select-Object -Property `
  @{Name='Student_id';Expression={[int]$PSItem.'Student ID'}},
  @{Name='Student_password';Expression={$PSItem.'Password'}} | Out-DataTable
  try {
    if (($passwords | Measure-Object).Count -ge 1) {
      Invoke-SQLiteBulkCopy -DataTable $passwords -DataSource $database -Table 'passwords' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
    }
  } catch {
    write-host "ERROR: Could not import student passwords correctly." -ForegroundColor RED
    exit(1)
  }
}

write-host "Info: Importing schools.csv."
$schools = import-csv .\files\schools.csv | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $schools -DataSource $database -Table 'schools' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import schools.csv correctly." -ForegroundColor RED
  exit(1)
}

#This is a cognos query and can have extra spaces on data. Forcing data type will fix most issues. FacultyIDs first then teachers.
write-host "Info: Importing facultyids.csv."
$facultyids = import-csv .\files\facultyids.csv | Select-Object -Property `
@{Name='School_id';Expression={[int]$PSItem.'School_id'}},
@{Name='Teacher_id';Expression={([string]$PSItem.'Teacher_id').Trim()}},
@{Name='Teacher_number';Expression={([string]$PSItem.'Teacher_id').Trim()}},
@{Name='State teacher id';Expression={[long]$PSItem.'State teacher id'}},
@{Name='Teacher_email';Expression={[string]$PSItem.'Teacher_email'}},
@{Name='First_name';Expression={[string]$PSItem.'First_name'}},
@{Name='Middle_name';Expression={''}},
@{Name='Last_name';Expression={[string]$PSItem.'Last_name'}},
@{Name='Title';Expression={''}},
@{Name='Username';Expression={''}},
@{Name='Password';Expression={''}},
'Timestamp'
$facultyids | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$facultyids = $facultyids | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $facultyids -DataSource $database -Table 'teachers' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import facultyids.csv correctly." -ForegroundColor RED
  exit(1)
}

write-host "Info: Importing teachers.csv."
$teachers = import-csv .\files\teachers.csv | Select-Object -Property *,Timestamp
$teachers | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$teachers = $teachers | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $teachers -DataSource $database -Table 'teachers' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import teachers.csv correctly." -ForegroundColor RED
  exit(1)
}

write-host "Info: Importing students.csv."
$students = import-csv .\files\students.csv | Select-Object -Property *,Timestamp
$students | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$students = $students | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $students -DataSource $database -Table 'students' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import students.csv correctly." -ForegroundColor RED
  exit(1)
}

#Because this is a query some properties come in with spaces. Forcing the data type will clear spaces around Student_id.
write-host "Info: Importing students_extras.csv."
$students_extras = Import-Csv .\files\students_extras.csv | Select-Object -Property `
  @{Name='Student_id';Expression={[int]$PSItem.'Student_id'}},
  @{Name='Student_gradyr';Expression={[int]$PSItem.'Student_gradyr'}},
  @{Name='Student_nickname';Expression={[string]($PSItem.'Student_nickname').Trim()}},
  @{Name='Student_homeroom';Expression={($PSItem.'Student_homeroom').Trim()}},
  @{Name='Student_hrmtid';Expression={($PSItem.'Student_hrmtid').Trim()}},
  @{Name='Student_advisor';Expression={($PSItem.'Student_advisor').Trim()}},
  @{Name='Student_houseteam';Expression={($PSItem.'Student_houseteam').Trim()}},
  @{Name='Student_haclogin';Expression={($PSItem.'Student_haclogin').Trim()}},
  @{Name='Student_hacpassword';Expression={$PSItem.'Student_hacpassword'}},
  @{Name='Student_contactid';Expression={[int]$PSItem.'Student_contactid'}},
  @{Name='Student_mealstatus';Expression={([string]$PSItem.'Student_mealstatus').Trim()}},
  'Timestamp'
  $students_extras | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
  $students_extras = $students_extras | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $students_extras -DataSource $database -Table 'students_extras' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import students_extras.csv correctly." -ForegroundColor RED
  exit(1)
}

write-host "Info: Importing contacts.csv."
$contacts = import-csv .\files\contacts.csv | Select-Object -Property `
  @{Name='Contact_id';Expression={[int]$PSItem.'Contact_id'}},
  @{Name='Student_id';Expression={[int]$PSItem.'Student_id'}},
  @{Name='Contact_priority';Expression={[int]$PSItem.'Contact_priority'}},
  @{Name='Contact_firstname';Expression={($PSItem.'Contact_firstname').Trim()}},
  @{Name='Contact_lastname';Expression={($PSItem.'Contact_lastname').Trim()}},
  @{Name='Contact_type';Expression={($PSItem.'Contact_type').Trim()}},
  @{Name='Contact_phonenumber';Expression={$PSItem.'Contact_phonenumber'}},
  @{Name='Contact_relationship';Expression={($PSItem.'Contact_relationship').Trim()}},
  @{Name='Contact_phonetype';Expression={($PSItem.'Contact_phonetype').Trim()}},
  @{Name='Contact_email';Expression={($PSItem.'Contact_email').Trim()}},
  'Timestamp'
$contacts | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$contacts = $contacts | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $contacts -DataSource $database -Table 'contacts' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import contacts.csv correctly." -ForegroundColor RED
  exit(1)
}

write-host "Info: Importing transportation.csv."
$transportation = import-csv .\files\transportation.csv | Select-Object -Property `
  @{Name='Student_id';Expression={[int]$PSItem.'Student_id'}},
  @{Name='Student_BusNumFrom';Expression={($PSItem.'Student_BusNumFrom').Trim()}},
  @{Name='Student_TravelTypeFrom';Expression={($PSItem.'Student_TravelTypeFrom').Trim()}},
  @{Name='Student_BusNumTo';Expression={($PSItem.'Student_BusNumTo').Trim()}},
  @{Name='Student_TravelTypeTo';Expression={($PSItem.'Student_TravelTypeTo').Trim()}},
  'Timestamp'
$transportation | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$transportation = $transportation | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $transportation -DataSource $database -Table 'transportation' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import transportation.csv correctly." -ForegroundColor RED
  exit(1)
}

write-host "Info: Importing sections.csv."
$sections = import-csv .\files\sections.csv | Select-Object -Property *,Timestamp | Where-Object { $PSitem.'Teacher_id' -ne 0 }
$sections | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$sections = $sections | Sort-Object -Property Section_id,Term_name | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $sections -DataSource $database -Table 'sections' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import sections.csv correctly." -ForegroundColor RED
  exit(1)
}

write-host "Info: Importing enrollments.csv."
$enrollments = import-csv .\files\enrollments.csv | Select-Object -Property *,Timestamp
$enrollments | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$enrollments = $enrollments | Sort-Object -Property Section_id,Marking_period | Out-DataTable
try {
  Invoke-SQLiteBulkCopy -DataTable $enrollments -DataSource $database -Table 'enrollments' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Could not import enrollments.csv correctly." -ForegroundColor RED
  exit(1)
}

#Because this is a query some properties come in with spaces. Forcing the data type will clear spaces around Student_id.
write-host "Info: Importing activities.csv."
$activities = Import-Csv .\files\activities.csv | Select-Object -Property `
  @{Name='School_id';Expression={[int]$PSItem.'School_id'}},
  @{Name='Student_id';Expression={[int]$PSItem.'Student_id'}},
  @{Name='Teacher_id';Expression={([string]$PSItem.'Teacher_id').Trim()}},
  @{Name='Activity_code';Expression={[string]$($PSItem.'Activity_code').Trim()}},
  @{Name='Activity_name';Expression={[string]$($PSItem.'Activity_name').Trim()}},
  'Timestamp' | Where-Object { $activitiesBuildings -contains $PSItem.'School_id' }
$includedActivites = @()
$activitiesInclude | ForEach-Object {
  $includeString = $PSItem
  $includedActivites += $activities | Where-Object { $PSItem.'Activity_name' -LIKE $includeString }
}
$activities = $includedActivites
$activitiesIgnore | ForEach-Object {
  $excludeString = $PSItem
  $activities = $activities | Where-Object { $PSItem.'Activity_name' -NOTLIKE $excludeString }
}
$activities | ForEach-Object { $PSItem.'Timestamp' = $timestamp }
$activities = $activities | Out-DataTable
try {
  if (($activities | Measure-Object).Count -ge 1) {
    Invoke-SQLiteBulkCopy -DataTable $activities -DataSource $database -Table 'activities' -Force -ConflictClause REPLACE -ErrorAction 'STOP'
  }
} catch {
  write-host "ERROR: Could not import activities.csv correctly." -ForegroundColor RED
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

#homerooms
if ($homeroomsScheduled) {
  write-host "Info: Building homerooms as a scheduled class."

  #delete students_extras no longer in CSV. We have to do this before enrollment is modified.
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM students_extras WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students_extras ORDER BY Timestamp DESC LIMIT 1);'
 
 
  $homeroomQuery = Invoke-SqliteQuery -Database $database -Query "select DISTINCT Students.School_id,Student_homeroom,Student_hrmtid,Teacher_email,teachers.First_name,teachers.Last_name `
    from students_extras `
    INNER JOIN students ON students_extras.Student_id = students.Student_id `
    INNER JOIN teachers ON students_extras.Student_hrmtid = teachers.Teacher_id `
    WHERE Student_homeroom != ''"

  $homerooms = @()
  $homeroomQuery | ForEach-Object {

    #School_id,Section_id,Teacher_id,Teacher_2_id,Teacher_3_id,Teacher_4_id,Name,Section_number,Grade,Course_name,Course_number,Course_description,Period,Subject,Term_name,Term_start,Term_end
    #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,1,08/13/2020,10/24/2020
    #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,2,08/13/2020,10/24/2020
    #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,3,08/13/2020,10/24/2020
    #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,4,08/13/2020,10/24/2020

    $homeroom = $PSItem.'Student_homeroom'
    $classid = "9999$($PSItem.'Student_homeroom')"
    $teacher = $PSItem.'Last_name'
    $schoolid = $PSItem.'School_id'

    #4x for each term.
    $homerooms += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Student_hrmtid'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$($teacher) Homeroom"; Section_number = 1; Grade = 1; Course_name = "$($teacher) Homeroom"; Course_number = ''; Course_description = "$($teacher) Homeroom"; Period = 0; Subject = ''; Term_name = 1; Term_start = $term1.'Term_start'; Term_end = $term1.'Term_end'; Timestamp = $timestamp }
    $homerooms += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Student_hrmtid'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$($teacher) Homeroom"; Section_number = 1; Grade = 1; Course_name = "$($teacher) Homeroom"; Course_number = ''; Course_description = "$($teacher) Homeroom"; Period = 0; Subject = ''; Term_name = 2; Term_start = $term2.'Term_start'; Term_end = $term2.'Term_end'; Timestamp = $timestamp }
    $homerooms += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Student_hrmtid'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$($teacher) Homeroom"; Section_number = 1; Grade = 1; Course_name = "$($teacher) Homeroom"; Course_number = ''; Course_description = "$($teacher) Homeroom"; Period = 0; Subject = ''; Term_name = 3; Term_start = $term3.'Term_start'; Term_end = $term3.'Term_end'; Timestamp = $timestamp }
    $homerooms += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Student_hrmtid'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$($teacher) Homeroom"; Section_number = 1; Grade = 1; Course_name = "$($teacher) Homeroom"; Course_number = ''; Course_description = "$($teacher) Homeroom"; Period = 0; Subject = ''; Term_name = 4; Term_start = $term4.'Term_start'; Term_end = $term4.'Term_end'; Timestamp = $timestamp }
    
    $studentids = Invoke-SqliteQuery -Database $database -Query "SELECT Student_id FROM students_extras WHERE Student_homeroom = $homeroom" | Select-Object -ExpandProperty Student_id
    
    $studentids | foreach-object {
        Invoke-SqliteQuery -Database $database -Query "INSERT OR REPLACE INTO enrollments (School_id, Section_id, Student_id, Marking_period, Timestamp) VALUES ($schoolid,$classid,$PSItem,1,$timestamp),($schoolid,$classid,$PSItem,2,$timestamp),($schoolid,$classid,$PSItem,3,$timestamp),($schoolid,$classid,$PSItem,4,$timestamp)"
    }
  
  }

  $homerooms = $homerooms | Out-DataTable
  try {
    Invoke-SQLiteBulkCopy -DataSource $database -Table sections -DataTable $homerooms -Force -ConflictClause REPLACE -ErrorAction 'STOP'
  } catch {
    write-host "ERROR: Could not import homerooms correctly." -ForegroundColor RED
    exit(1)
  }
}

#activities
if ($activitiesScheduled) {
  write-host "Info: Building activities as a scheduled class."

  #delete activities no longer in CSV. We have to do this before enrollment is modified.
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM activities WHERE Timestamp != (SELECT DISTINCT Timestamp FROM activities ORDER BY Timestamp DESC LIMIT 1);'
 
  $activitiesQuery = Invoke-SqliteQuery -Database $database -Query "SELECT DISTINCT School_id,Teacher_id,Activity_code,Activity_name from activities"
  if ($null -eq $activitiesQuery) {
    Write-Host "Error: You have specified in the settings.ps1 that you would like to build activities as a scheduled class. However, no activities were found in the database." -ForegroundColor RED
  } else {

    $activities = @()
    $activitiesQuery | ForEach-Object {

      #Need to add the school name to the front of the Activity Code bc of duplicate activity names between buildings.
      $schoolid = $PSItem.School_id
      $actPrefix = $validschools.[int]$schoolid
      $classid = ([string](Get-FNVHash "$($actPrefix)$($PSItem.'Activity_code')")).PadLeft(16,"9") #a deterministic way to end up with a positive integer based on the name. Then pad with 9's to 16 characters to avoid conflicts.
      $activityName = $PSItem.'Activity_name'

      #School_id,Section_id,Teacher_id,Teacher_2_id,Teacher_3_id,Teacher_4_id,Name,Section_number,Grade,Course_name,Course_number,Course_description,Period,Subject,Term_name,Term_start,Term_end
      #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,1,08/13/2020,10/24/2020
      #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,2,08/13/2020,10/24/2020
      #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,3,08/13/2020,10/24/2020
      #15,999504,1234,,,,cmillsap homeroom,1,1,cmillsap homeroom,,cmillsap-homeroom,0,,4,08/13/2020,10/24/2020

      #4x for each term.
      $activities += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Teacher_id'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$activityName"; Section_number = 1; Grade = ''; Course_name = "$activityName"; Course_number = ''; Course_description = "$activityName"; Period = 99; Subject = ''; Term_name = 1; Term_start = $term1.'Term_start'; Term_end = $term1.'Term_end'; Timestamp = $timestamp }
      $activities += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Teacher_id'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$activityName"; Section_number = 1; Grade = ''; Course_name = "$activityName"; Course_number = ''; Course_description = "$activityName"; Period = 99; Subject = ''; Term_name = 2; Term_start = $term2.'Term_start'; Term_end = $term2.'Term_end'; Timestamp = $timestamp }
      $activities += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Teacher_id'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$activityName"; Section_number = 1; Grade = ''; Course_name = "$activityName"; Course_number = ''; Course_description = "$activityName"; Period = 99; Subject = ''; Term_name = 3; Term_start = $term3.'Term_start'; Term_end = $term3.'Term_end'; Timestamp = $timestamp }
      $activities += [PSCustomObject]@{School_id = "$($PSItem.'School_id')"; Section_id = $classid; Teacher_id = $($PSItem.'Teacher_id'); Teacher_2_id = ''; Teacher_3_id = ''; Teacher_4_id = ''; Name = "$activityName"; Section_number = 1; Grade = ''; Course_name = "$activityName"; Course_number = ''; Course_description = "$activityName"; Period = 99; Subject = ''; Term_name = 4; Term_start = $term4.'Term_start'; Term_end = $term4.'Term_end'; Timestamp = $timestamp }
      
      #Pull student activites that are in their own building ONLY!
      $studentids = Invoke-SqliteQuery -Database $database -Query "SELECT activities.Student_id FROM activities INNER JOIN students ON students.Student_id = activities.Student_id WHERE students.School_id = $($PSItem.'School_id') AND activities.School_id = $($PSItem.'School_id') AND Activity_code = ""$($PSItem.'Activity_code')""" | Select-Object -ExpandProperty Student_id
      
      $studentids | foreach-object {
        Invoke-SqliteQuery -Database $database -Query "INSERT OR REPLACE INTO enrollments (School_id, Section_id, Student_id, Marking_period, Timestamp) VALUES ($schoolid,$classid,$PSItem,1,$timestamp),($schoolid,$classid,$PSItem,2,$timestamp),($schoolid,$classid,$PSItem,3,$timestamp),($schoolid,$classid,$PSItem,4,$timestamp)"
      }
    
    }

    $activities = $activities | Out-DataTable
    try {
      Invoke-SQLiteBulkCopy -DataSource $database -Table sections -DataTable $activities -Force -ConflictClause REPLACE -ErrorAction 'STOP'
    } catch {
      write-host "ERROR: Could not import activities correctly." -ForegroundColor RED
      exit(1)
    }
  }
}

###########################################################
# Remove old teachers, students, sections, enrollments
###########################################################
if (-Not($SkipTableCleanup)) {
  write-host "Info: Archiving and deleting old entries. DOES NOT INCLUDE students_extras OR activities tables!"
  #Archive students not in CSV
  $q = 'INSERT INTO students_disabled SELECT * FROM students WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1);
  DELETE FROM students WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1);'
  Invoke-SqliteQuery -DataSource $database -Query $q

  #delete teachers no longer in CSV
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM teachers WHERE Timestamp != (SELECT DISTINCT Timestamp FROM teachers ORDER BY Timestamp DESC LIMIT 1);'

  #delete contacts no longer in CSV
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM contacts WHERE Timestamp != (SELECT DISTINCT Timestamp FROM contacts ORDER BY Timestamp DESC LIMIT 1);'

  #delete transportation no longer in CSV
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM transportation WHERE Timestamp != (SELECT DISTINCT Timestamp FROM transportation ORDER BY Timestamp DESC LIMIT 1);'

  #remove invalid buildings that might pull in with the facultyids.
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM teachers WHERE School_id NOT IN (SELECT School_id from schools)'

  #delete sections no longer in CSV
  Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM sections WHERE Timestamp != (SELECT DISTINCT Timestamp FROM sections ORDER BY Timestamp DESC LIMIT 1);'

  #archived and delete enrollments no longer in CSV
  if ($RetainEnrollmentsDays -ge 1) {
    #Retain enrollments for the given time frame. Archive to the enrollments_archived after.
    Invoke-SqliteQuery -DataSource $database -Query "INSERT OR REPLACE INTO enrollments_archived SELECT * FROM enrollments WHERE Timestamp < $timestampRetain;"
    #Just in case this setting is changed we need to pull back in from the enrollments_archived table.
    Invoke-SqliteQuery -DataSource $database -Query "INSERT OR REPLACE INTO enrollments SELECT * FROM enrollments_archived WHERE Timestamp > $timestampRetain;"
    #Clean up older enrollments from enrollments table.
    Invoke-SqliteQuery -DataSource $database -Query "DELETE FROM enrollments WHERE Timestamp < $timestampRetain;"
  } else {
    #Retain enrollments data by copying to the enrollments_archived table.
    Invoke-SqliteQuery -DataSource $database -Query 'INSERT OR REPLACE INTO enrollments_archived SELECT * FROM enrollments WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1);'
    Invoke-SqliteQuery -DataSource $database -Query 'DELETE FROM enrollments WHERE Timestamp != (SELECT DISTINCT Timestamp FROM enrollments ORDER BY Timestamp DESC LIMIT 1);'
  }

}
######################################################################################################################
# Group Enrollments and Sections together to match for an actual Schedule.
######################################################################################################################

write-host "Info: Building Merged Sections and Enrollments to build correct Schedules."

#build schedules based on matching enrollments
$q = '/* FINAL REVISION! */
/* To ensure the table columns are correct I am dropping the entire table instead of truncating. */
DROP TABLE IF EXISTS enrollments_grouped;
CREATE TABLE "enrollments_grouped" (
	"School_id"	INTEGER,
	"Section_id"	TEXT,
	"Student_id"	INTEGER,
	"Terms"	INTEGER
);

/* I need the terms to be grouped together to query later. Make a copy of the enrollments table with all terms grouped together. */
INSERT INTO enrollments_grouped 
  SELECT enrollments.School_id,Section_id,enrollments.Student_id,group_concat(Marking_period,'''') as Terms
  FROM enrollments
  LEFT JOIN students ON enrollments.Student_id = students.Student_id
  GROUP BY enrollments.Student_id,enrollments.Section_id;

/* I need the terms to be grouped together to query later. Make a copy of the sections table with all terms grouped together. */
DROP TABLE IF EXISTS sections_grouped;
CREATE TABLE sections_grouped(
  Terms INT,
  School_id INT,
  Section_id INT,
  Teacher_id TEXT,
  Name TEXT,
  Section_number INT,
  Grade INT,
  Course_name TEXT,
  Course_number TEXT,
  Course_description TEXT,
  Period TEXT,
  Subject TEXT,
  Term_start TEXT,
  Term_end TEXT
);

INSERT INTO sections_grouped
  SELECT group_concat(Term_name,'''') as Terms,School_id, Section_id, Teacher_id, Name, Section_number, Grade, Course_name, Course_number, Course_description, Period, Subject, Term_start, Term_end
  FROM sections
  GROUP BY sections.Section_id;

/*This table will hold our final results. Matching all kids who are enrolled for the entire length of the section terms. Leaving out classes that do not have a teacher email assigned to them. */
DROP TABLE IF EXISTS schedules;
CREATE TABLE schedules(
  School_id INT,
  Section_id TEXT,
  Course_number TEXT,
  Section_number INT,
  Terms INT,
  Student_id INT,
  Teacher_id TEXT,
  Period TEXT,
  First_name TEXT,
  Last_name TEXT,
  Student_email TEXT collate nocase,
  Teacher_firstname TEXT,
  Teacher_lastname TEXT,
  Teacher_email TEXT collate nocase,
  UNIQUE(School_id,Section_id,Student_id)
);

/* This requires that teachers have an email address. This is on purpsoe for my district because of sections that are placeholders. */
INSERT INTO schedules
  SELECT enrollments_grouped.School_id, enrollments_grouped.Section_id, sections_grouped.Course_number, sections_grouped.Section_number, enrollments_grouped.Terms, enrollments_grouped.Student_id, sections_grouped.Teacher_id, sections_grouped.Period, students.First_name, students.Last_name, students.Student_email, teachers.First_name as Teacher_firstname, teachers.Last_name as Teacher_lastname, teachers.Teacher_email
  FROM enrollments_grouped
  INNER JOIN sections_grouped ON enrollments_grouped.Section_id = sections_grouped.Section_id AND enrollments_grouped.Terms = sections_grouped.Terms
  INNER JOIN students ON students.Student_id = enrollments_grouped.Student_id
  INNER JOIN teachers ON sections_grouped.Teacher_id = teachers.Teacher_id AND enrollments_grouped.School_id = teachers.School_id
  WHERE Teacher_email != '''' ORDER BY enrollments_grouped.Student_id,sections_grouped.Period;
'

try {
  Invoke-SqliteQuery -DataSource $database -Query $q -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Failed to create merged sections and enrollments!" -ForegroundColor RED
  exit(1)
}

$q = "/* Now insert just the current Term to fix enrollment. We ignore duplicates */
INSERT OR IGNORE INTO schedules SELECT enrollments.School_id,enrollments.Section_id, sections.Course_number, sections.Section_number, enrollments.Marking_period as 'Terms', enrollments.Student_id, sections.Teacher_id, sections.Period, students.First_name, students.Last_name, students.Student_email, teachers.First_name as Teacher_firstname, teachers.Last_name as Teacher_lastname, teachers.Teacher_email FROM enrollments INNER JOIN sections ON enrollments.Section_id = sections.Section_id INNER JOIN students ON students.Student_id = enrollments.Student_id INNER JOIN teachers ON sections.Teacher_id = teachers.Teacher_id AND enrollments.School_id = teachers.School_id WHERE Teacher_email != '' AND Marking_period = $currentTerm
ORDER BY enrollments.Student_id,sections.Period;
"
try {
  Invoke-SqliteQuery -DataSource $database -Query $q -ErrorAction 'STOP'
} catch {
  write-host "ERROR: Failed to create schedules based on merged sections and enrollments!" -ForegroundColor RED
  exit(1)
}

#Verify we have content in the grouped tables and schedules. We have to know we have data in there.
try {
  $count = Invoke-SqliteQuery -DataSource $database -Query "SELECT COUNT(*) AS count FROM students" -ErrorAction 'STOP' | Select-Object -ExpandProperty count
  if ($count -le $minStuCount) { write-host "ERROR: Not enough students to pass verification. Please double check your import file for a students." -ForegroundColor RED; exit(1) }
  $count = Invoke-SqliteQuery -DataSource $database -Query "SELECT COUNT(*) AS count FROM sections_grouped" -ErrorAction 'STOP' | Select-Object -ExpandProperty count
  if ($count -le $minStuCount) { write-host "ERROR: Not enough sections_grouped to pass verification. Please double check your import file for a sections." -ForegroundColor RED; exit(1) }
  $count = Invoke-SqliteQuery -DataSource $database -Query "SELECT COUNT(*) AS count FROM enrollments_grouped" -ErrorAction 'STOP' | Select-Object -ExpandProperty count
  if ($count -le $minStuCount) { write-host "ERROR: Not enough enrollments_grouped to pass verification. Please double check your import file for a enrollments." -ForegroundColor RED; exit(1) }
  $count = Invoke-SqliteQuery -DataSource $database -Query "SELECT COUNT(*) AS count FROM schedules" -ErrorAction 'STOP' | Select-Object -ExpandProperty count
  if ($count -le $minStuCount) { write-host "ERROR: Not enough schedules to pass verification. Please double check your import file for a schedules." -ForegroundColor RED; exit(1) }
} catch {
  write-host "ERROR: Failed to verify count in students, grouped sections, grouped enrollments, and/or schedules." -ForegroundColor RED
  exit(1)
}


# Final Script to do all the stuff.
if (-Not($DisablePostProcessingScript)) {
  if (-Not($Staging)) {
      if (Test-Path $currentPath\x_PostProcessingSQLite.ps1) {
          . $currentPath\x_PostProcessingSQLite.ps1
      }
  } else {
      Write-Host "Info: Staging specified. Will not run x_PostProcessingSQLite.ps1."
  }
} else {
  Write-Host "Info: DisablePostProcessingScript specified. Will not run x_PostProcessingSQLite.ps1."
}

Stop-TranScript