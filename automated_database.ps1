#Requires -Version 7.0
#Requires -Modules SimplySQL,CognosModule

#Get-Help .\automated_database.ps1
#Get-Help .\automated_database.ps1 -Examples
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
  This script is used to download required Cognos reports to build a local sqlite database that we can pull custom data reports from.
  
  .NOTES
    Author: Craig Millsap
    Creation Date: 3/2021

  .EXAMPLE
  PS> .\automated_database.ps1 -DisablePostProcessingScript
  This will download the required Cognos Files and build the local database but it will not move to the next step by invoking x_PostProcessingDatabase.ps1.

  .EXAMPLE
  PS> .\automated_database.ps1 -SkipDownloadingReports
  Its possible to provide your own files and build your database from the CSVs in .\files\.

  .EXAMPLE
  PS> .\automated_database.ps1 -QuickMode
  This will only download the students.csv and students_extras.csv. It will skip over sections, enrollments, ect. Note: we can't create export files without the additional data.
  
  .PARAMETER DownloadFilesOnly
  Download all Cognos Reports but do not continue.

  .PARAMETER ClearDatabase
  A quick way to start your database over without losing archived data.
  
  .PARAMETER SkipTableCleanup
  This will not remove rows that do not have the current timestamp. This allows you to retain all previous data without anything expiring from the database.
  
  .PARAMETER Term
  This was added as a way to switch enrollment data into a different term. This was needed after some major schedule changes. Using the $RetainEnrollmentsDays in settings.ps1 allowed to an overlap of enrollment data.

  .PARAMETER QuickMode
  This mode will only download students.csv and students_extras.csv and start the process of creating student accounts. It calls automated_students.ps1 -QuickMode.
  This will create student accounts for only mismatched on Student ID, First Name, and Last Name. It will loop back around and build the complete database to move forward to file exports.

#>

Param(
    [Parameter(Mandatory=$false)][switch]$ClearDatabase, #Drops all tables except for archived so it can be recreated.
    [Parameter(Mandatory=$false)][switch]$SkipDownloadingReports, #Do not download updated report files using the CognosModule.
    [Parameter(Mandatory=$false)][switch]$SkipSanitizingFiles, #Do not try and remove special characters from files before processing.
    [Parameter(Mandatory=$false)][switch]$DownloadFilesOnly, #Download new files from Cognos then exit.
    [Parameter(Mandatory=$false)][switch]$SkipTableCleanup, #Do not remove rows that doen't match current timestamp.
    [Parameter(Mandatory=$false)][switch]$DisablePostProcessingScript, #Do not run the final script.
    [Parameter(Mandatory=$false)][ValidateSet(1,2,3,4)][int]$Term, #Specify the term to process for Schedules. This can be used to revert back a Clever export.
    [Parameter(Mandatory=$false)][switch]$QuickMode #This will only import the Students and Students_extras reports.
)

$scriptVersion = 1.3
$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

if (-Not(Test-Path $currentPath\logs)) { New-Item -ItemType Directory -Path $currentPath\logs }
$logfile = "$currentPath\logs\automated_database_$(get-date -f yyyy-MM-dd-HH-mm-ss).log"
try {
    Start-Transcript($logfile)
} catch {
    Stop-TranScript; Start-Transcript($logfile)
}

#SimplySQL Module is required.
if (Get-Module -ListAvailable | Where-Object {$PSItem.name -eq "SimplySQL"}) {
  Try { Import-Module SimplySQL } catch { Write-Host "Error: Unable to load module SimplySQL." -ForegroundColor RED; exit(1) }
} else {
  Write-Host 'SimplySQL Module not found!'
  if ($(@('Y','y','YES','yes','Yes')) -contains $(Read-Host -Prompt 'Would you like to try and automatically install it? y/n')) {
      try {
          Install-Module -Name SimplySQL -Scope AllUsers -Force
          Import-Module SimplySQL
      } catch {
          write-host 'Failed to install SimplySQL Module.' -ForegroundColor RED
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

#Attempt connecting to the database.
try {
  Connect-Database -Database $database
} catch {
  Write-Host "Error: Failed to connect to database engine." -ForegroundColor Red; exit(1)
}

$timestamp = [int64](Get-Date -UFormat %s)
if ($RetainEnrollmentsDays -ge 1) {
  $timestampRetain = [long] (Get-Date -Date (((Get-Date).AddDays(-$RetainEnrollmentsDays)).ToUniversalTime()) -UFormat %s)
}
if ($RetainStudentsDays -ge 1) {
  $timestampRetainStudents = [long] (Get-Date -Date (((Get-Date).AddDays(-$RetainStudentsDays)).ToUniversalTime()) -UFormat %s)
}
$timestamp365DaysAgo = [long] (Get-Date -Date (((Get-Date).AddYears(-1)).ToUniversalTime()) -UFormat %s)

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

#Quick way to clear out database.
if ($ClearDatabase) {
  write-host "Info: Clear database was specified. Dropping all tables except for disabled_students and enrollments_archived. These are needed to retain enrollments." -Foregroundcolor YELLOW
  Invoke-SqlUpdate -Query 'drop table if exists schools;
    drop table if exists teachers;
    drop table if exists teachers_extras;
    drop table if exists staff;
    drop table if exists students;
    drop table if exists sections;
    drop table if exists enrollments;
    drop table if exists sections_grouped;
    drop table if exists enrollments_grouped;
    drop table if exists schedules;
    drop table if exists students_extras;
    drop table if exists contacts;
    drop table if exists activities;
    drop table if exists transportation;
    drop table if exists hac_guardians;
    drop table if exists hac_students;'
  exit
}

if ($QuickMode) {
  $reports = @{
      'students' = @{ parameters = ''; folder = 'automation'; arguments = '' }
      'students_extras' = @{ parameters = ''; folder = 'automation'; arguments = '' }
  }
} else {
  $reports = @{
    'enrollments' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'schools' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'sections' = @{ parameters = "p_year=$([string]$schoolyear)"; folder = 'automation'; arguments = '' }
    'students' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'teachers' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'teachers_extras' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'staff' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'students_extras' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'contacts' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'activities' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'transportation' = @{ parameters = ''; folder = 'automation'; arguments = '' }
    'hac_students' = @{ parameters = ''; folder = 'hac_students'; arguments = '' }
    'hac_guardians' = @{ parameters = ''; folder = 'hac_guardians'; arguments = '' }
  }
}

if ($SkipDownloadingReports -eq $False) {

  #Establish Session Only. Report parameter is required but we can provide a fake one for authentication only.
  try {
    Connect-ToCognos
  } catch {
    Send-EmailNotification -subject "Automated_Students: Unable to establish connection to Cognos." -Body "Automated_students was unable to login to Cognos Analytics."
    exit(1)
  }

  $reports.keys | ForEach-Object {

    $report = ($reports).$PSItem

    write-host "INFO: Downloading $PSItem"
    
    $CognosArgs = @{
      report = $PSItem
      savepath = "$($currentPath)\files\"
      reportparams = (($reports).$PSItem).parameters
      cognosfolder = "_Shared Data File Reports\automated_students\21.4.8"
    }

    try {
      Save-CognosReport @CognosArgs -TeamContent -TrimCSVWhiteSpace
    } catch {
      #Write-Output can be used to Receive-Job
      Write-Output "$($CognosArgs.report): $PSitem"
      Send-EmailNotification -subject "Automated_Students: Failed to download Cognos Report." -Body "Automated_students was unable to download $($CognosArgs.report)"
      exit(1)
    }

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

if ($DownloadFilesOnly) { exit }

# Sanitize files by removing any special characters
if (-Not($SkipSanitizingFiles)) {
  write-host "Info: Sanitizing and removing special characters from import files."
  $reports.Keys | ForEach-Object {
      $fileContents = Get-Content -Encoding utf8 ".\files\$($PSItem).csv" | Remove-StringLatinCharacter
      $fileContents | Set-Content -Encoding utf8 ".\files\$($PSItem).csv"
      $fileContents = $NULL
  }
}

#Verify that the files imported have at least a School_id or Student_id property and a count greater than 3. I don't know any schools with less than 3 campuses. #activites.csv and teacher_extras is an exception.
write-host "Info: Verifying each file has a School_id or Student_id column."
$reports.Keys | ForEach-Object {
  if (@('activities','teachers_extras') -notcontains $PSItem) {
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

Start-SqlTransaction

###########################################################
#Verify tables exist and create if not.
###########################################################
write-host "Info: Check if database tables exist and create if not."

Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `db_settings` (
  `Name` varchar(32),
  `Setting` varchar(256),
  PRIMARY KEY (`Name`)
  );' | Out-Null

#Schools Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `schools` (
  `School_id` int(4),
  `School_name` varchar(64),
  `School_number` varchar(12),
  `State_id` varchar(12),
  `Low_grade` varchar(15),
  `High_grade` varchar(2),
  `Principal` varchar(64),
  `Principal_email` varchar(128),
  `School_address` varchar(64),
  `School_city` varchar(64),
  `School_state` varchar(2),
  `School_zip` int(5),
  `School_phone` bigint(10),
  `Timestamp` int(10),
  PRIMARY KEY (`School_id`)
  );' | Out-Null


#Teachers Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `teachers` (
  `School_id` int(4),
  `Teacher_id` varchar(32),
  `Teacher_number` varchar(32),
  `State_teacher_id` bigint(12),
  `Teacher_email` varchar(128),
  `First_name` varchar(64),
  `Middle_name` varchar(64),
  `Last_name` varchar(64),
  `Title` varchar(64),
  `Username` varchar(64),
  `Password` varchar(64),
  `Timestamp` int(10),
  UNIQUE (`School_id`,`Teacher_id`)
  );' | Out-Null

#Teachers_extras Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `teachers_extras` (
  `Teacher_id` varchar(32),
  `Street_name` varchar(128),
  `Complex` varchar(128),
  `State_id` bigint(12),
  `Login_id` varchar(32),
  `Timestamp` int(10),
  PRIMARY KEY (`Teacher_id`)
  );' | Out-Null

#Staff Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `staff` (
  `School_id` int(4),
  `Staff_id` varchar(32),
  `Staff_email` varchar(128),
  `First_name` varchar(64),
  `Last_name` varchar(64),
  `Department` varchar(64),
  `Title` varchar(64),
  `Username` varchar(128),
  `Password` varchar(128),
  `Role` varchar(64),
  `Is_Advisor` varchar(1),
  `Is_Counselor` varchar(1),
  `Is_Teacher` varchar(1),
  `Timestamp` int(10),
  UNIQUE (`School_id`,`Staff_id`)
  );' | Out-Null

#Students Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `students` (
  `School_id` int(4),
  `Student_id` int(10),
  `Student_number` int(10),
  `State_id` bigint(12),
  `Last_name` varchar(64),
  `Middle_name` varchar(1),
  `First_name` varchar(64),
  `Grade` varchar(16),
  `Gender` varchar(1),
  `DOB` varchar(10),
  `Race` varchar(1),
  `Hispanic_Latino` varchar(1),
  `Ell_status` varchar(1),
  `Frl_status` varchar(1),
  `Iep_status` varchar(1),
  `Student_street` varchar(64),
  `Student_city` varchar(32),
  `Student_state` varchar(2),
  `Student_zip` varchar(10),
  `Student_email` varchar(128),
  `Contact_relationship` varchar(19),
  `Contact_type` varchar(8),
  `Contact_name` varchar(64),
  `Contact_phone` varchar(13),
  `Contact_email` varchar(128),
  `Username` varchar(10),
  `Password` varchar(10),
  `Home_language` varchar(25),
  `Timestamp` int(10),
  PRIMARY KEY (`Student_id`)
  );' | Out-Null

#Students_disabled Table - CREATE TABLE AS not implimented in PSSQLite module.
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `students_disabled` (
  `School_id` int(4),
  `Student_id` int(10),
  `Student_number` int(9),
  `State_id` bigint(12),
  `Last_name` varchar(18),
  `Middle_name` varchar(1),
  `First_name` varchar(12),
  `Grade` varchar(12),
  `Gender` varchar(1),
  `DOB` varchar(10),
  `Race` varchar(1),
  `Hispanic_Latino` varchar(1),
  `Ell_status` varchar(1),
  `Frl_status` varchar(1),
  `Iep_status` varchar(1),
  `Student_street` varchar(37),
  `Student_city` varchar(15),
  `Student_state` varchar(2),
  `Student_zip` int(5),
  `Student_email` varchar(44),
  `Contact_relationship` varchar(19),
  `Contact_type` varchar(8),
  `Contact_name` varchar(24),
  `Contact_phone` varchar(13),
  `Contact_email` varchar(35),
  `Username` varchar(10),
  `Password` varchar(10),
  `Home_language` varchar(25),
  `Timestamp` int(10)
  );' | Out-Null

#Sections Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `sections` (
  `School_id` int(4),
  `Section_id` bigint(32),
  `Teacher_id` varchar(32),
  `Teacher_2_id` varchar(32),
  `Teacher_3_id` varchar(32),
  `Teacher_4_id` varchar(32),
  `Name` varchar(64),
  `Section_number` int(12),
  `Grade` varchar(15),
  `Course_name` varchar(64),
  `Course_number` varchar(32),
  `Course_description` varchar(64),
  `Period` varchar(4),
  `Subject` varchar(32),
  `Term_name` int(1),
  `Term_start` varchar(10),
  `Term_end` varchar(10),
  `Timestamp` int(12),
  UNIQUE (`School_id`,`Section_id`,`Term_name`)
  );' | Out-Null

#Enrollments Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `enrollments` (
  `School_id` int(4),
  `Section_id` bigint(32),
  `Student_id` int(10),
  `Marking_period` int(1),
  `Timestamp` int(12),
  UNIQUE (`School_id`,`Section_id`,`Student_id`,`Marking_period`)
 )' | Out-Null

#Enrollments_archived Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `enrollments_archived` (
  `School_id` int(4),
  `Section_id` bigint(32),
  `Student_id` int(10),
  `Marking_period` int(1),
  `Timestamp` int(12),
  UNIQUE (`School_id`,`Section_id`,`Student_id`,`Marking_period`)
 )' | Out-Null

#Passwords Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `passwords` (
  `Student_id` int(10),
  `Student_password` varchar(10),
  `HAC_passwordset` int(1),
  `Timestamp` int(12),
  UNIQUE (`Student_id`)
  );' | Out-Null

#students_extras Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `students_extras` (
  `Student_id` int(10) NOT NULL,
  `Student_gradyr` int(4) DEFAULT NULL,
  `Student_nickname` varchar(64) DEFAULT NULL,
  `Student_homeroom` varchar(10) DEFAULT NULL,
  `Student_hrmtid` varchar(12) DEFAULT NULL,
  `Student_advisor` varchar(12) DEFAULT NULL,
  `Student_houseteam` varchar(12) DEFAULT NULL,
  `Student_haclogin` varchar(128) DEFAULT NULL,
  `Student_hacpassword` varchar(256) DEFAULT NULL,
  `Student_contactid` int(12) DEFAULT NULL,
  `Student_mealstatus` varchar(2) DEFAULT NULL,
  `Timestamp` int(12),
  PRIMARY KEY (`Student_id`)
  )' | Out-Null

#Contacts Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `contacts` (
  `Contact_id` int(5) DEFAULT NULL,
  `Student_id` int(10) DEFAULT NULL,
  `Contact_priority` int(1) DEFAULT NULL,
  `Contact_firstname` varchar(20) DEFAULT NULL,
  `Contact_lastname` varchar(21) DEFAULT NULL,
  `Contact_type` varchar(1) DEFAULT NULL,
  `Contact_phonenumber` varchar(10) DEFAULT NULL,
  `Contact_relationship` varchar(19) DEFAULT NULL,
  `Contact_phonetype` varchar(2) DEFAULT NULL,
  `Contact_email` varchar(35) DEFAULT NULL,
  `Timestamp` int(12),
  UNIQUE (`Contact_id`,`Student_id`,`Contact_phonetype`)
  );' | Out-Null

#hac_students Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `hac_students` (
  `Contact_id` int(5) DEFAULT NULL,
  `Student_id` int(10) DEFAULT NULL,
  `Student_haclogin` varchar(44) DEFAULT NULL,
  `Student_webaccess` varchar(1) DEFAULT NULL,
  `Student__lastpasswordchange` varchar(19) DEFAULT NULL,
  `Student_lastlogindate` varchar(19) DEFAULT NULL,
  `Student_accesstoken` varchar(36) DEFAULT NULL,
  `Student_accesstokenused` varchar(19) DEFAULT NULL,
  `Timestamp` int(12),
  UNIQUE (`Contact_id`,`Student_id`)
  )' | Out-Null

#hac_guardians Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `hac_guardians` (
  `Contact_id` int(5) DEFAULT NULL,
  `Student_id` int(10) DEFAULT NULL,
  `Guardian_firstname` varchar(20) DEFAULT NULL,
  `Guardian_lastname` varchar(21) DEFAULT NULL,
  `Guardian_priority` int(1) DEFAULT NULL,
  `Guardian_haclogin` varchar(35) DEFAULT NULL,
  `Guardian_webaccess` varchar(1) DEFAULT NULL,
  `Guardian_passchangenextlogin` varchar(1) DEFAULT NULL,
  `Guardian_lastpasswordchange` varchar(19) DEFAULT NULL,
  `Guardian_lastlogindate` varchar(19) DEFAULT NULL,
  `Guardian_accesstoken` varchar(36) DEFAULT NULL,
  `Guardian_accesstokenused` varchar(19) DEFAULT NULL,
  `Timestamp` int(12),
  UNIQUE (`Contact_id`,`Student_id`)
  )' | Out-Null

#Activities Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `activities` (
  `School_id` int(4) DEFAULT NULL,
  `Student_id` int(10) DEFAULT NULL,
  `Teacher_id` varchar(32) DEFAULT NULL,
  `Activity_code` varchar(5) DEFAULT NULL,
  `Activity_name` varchar(28) DEFAULT NULL,
  `Timestamp` int(12),
  UNIQUE (`School_id`,`Student_id`,`Teacher_id`,`Activity_code`)
 );' | Out-Null

#Transportation Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `transportation` (
  `Student_id` int(10),
  `Student_BusNumFrom` varchar(2) DEFAULT NULL,
  `Student_TravelTypeFrom` varchar(1) DEFAULT NULL,
  `Student_BusNumTo` varchar(2) DEFAULT NULL,
  `Student_TravelTypeTo` varchar(1) DEFAULT NULL,
  `Timestamp` int(12),
  PRIMARY KEY (`Student_id`)
  );' | Out-Null

#Transportation Table
Invoke-SqlUpdate -Query 'CREATE TABLE IF NOT EXISTS `action_log` (
  `Student_id` int(10),
  `Action` varchar(64),
  `Identity` varchar(128),
  `Timestamp` varchar(19)
  );' | Out-Null

#`sections_Grouped`,`enrollments_Grouped`, and `schedules` tables are dropped/created/imported at the end of this script.

Complete-SqlTransaction
Start-SqlTransaction

###########################################################
# Import CSVs into Database Tables
###########################################################
write-host "Info: Importing CSV files into tables."

write-host "Info: Importing students.csv."
Import-CSV .\files\students.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | Where-Object { $ExcludeGrades -notcontains $PSitem.Grade } | ForEach-Object {
  $student = @{}
  $PSItem.psobject.properties | ForEach-Object {
    if ($_.Name -eq 'State_id') { $_.Value = [int64]$_.Value }
    $student[$_.Name] = $_.Value
  }
  #$student.GetType()
  try {
    $query = '
    (`School_id`,`Student_id`,`Student_number`,`State_id`,`Last_name`,`Middle_name`,`First_name`,`Grade`,`Gender`,`DOB`,`Race`,`Hispanic_Latino`,`Ell_status`,`Frl_status`,`Iep_status`,`Student_street`,`Student_city`,`Student_state`,`Student_zip`,`Student_email`,`Contact_relationship`,`Contact_type`,`Contact_name`,`Contact_phone`,`Contact_email`,`Username`,`Password`,`Home_language`,`Timestamp`)
    VALUES
    (@School_id,@Student_id,@Student_number,@State_id,@Last_name,@Middle_name,@First_name,@Grade,@Gender,@DOB,@Race,@Hispanic_Latino,@Ell_status,@Frl_status,@Iep_status,@Student_street,@Student_city,@Student_state,@Student_zip,@Student_email,@Contact_relationship,@Contact_type,@Contact_name,@Contact_phone,@Contact_email,@Username,@Password,@Home_language,@Timestamp)'
    
    if (@('mysql','mariadb') -contains $database.dbtype) {
      $query = 'INSERT INTO `students`',$query,'ON DUPLICATE KEY UPDATE `School_id`=@School_id, `Student_id`=@Student_id, `Student_number`=@Student_number, `State_id`=@State_id, `Last_name`=@Last_name, `Middle_name`=@Middle_name, `First_name`=@First_name, `Grade`=@Grade, `Gender`=@Gender, `DOB`=@DOB, `Race`=@Race, `Hispanic_Latino`=@Hispanic_Latino, `Ell_status`=@Ell_status, `Frl_status`=@Frl_status, `Iep_status`=@Iep_status, `Student_street`=@Student_street, `Student_city`=@Student_city, `Student_state`=@Student_state, `Student_zip`=@Student_zip, `Student_email`=@Student_email, `Contact_relationship`=@Contact_relationship, `Contact_type`=@Contact_type, `Contact_name`=@Contact_name, `Contact_phone`=@Contact_phone, `Contact_email`=@Contact_email, `Username`=@Username, `Password`=@Password, `Home_language`=@Home_language, `Timestamp`=@Timestamp'
      Invoke-SqlUpdate -Query $query -Parameters $student | Out-Null
    } else {
      $query = 'INSERT OR REPLACE INTO `students`',$query
      Invoke-SqlUpdate -Query $query -Parameters $student | Out-Null
    }
  } catch {
    write-host "Error: $PSItem"
    exit(1)
  }
}

#Because this is a query some properties come in with spaces. Forcing the data type will clear spaces around Student_id.
write-host "Info: Importing students_extras.csv."
Import-CSV .\files\students_extras.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
  $hashtable = @{}
  $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
  try {
    $query = '(`Student_id`,`Student_gradyr`,`Student_nickname`,`Student_homeroom`,`Student_hrmtid`,`Student_advisor`,`Student_houseteam`,`Student_haclogin`,`Student_hacpassword`,`Student_contactid`,`Student_mealstatus`,`Timestamp`)
    VALUES
    (@Student_id,@Student_gradyr,@Student_nickname,@Student_homeroom,@Student_hrmtid,@Student_advisor,@Student_houseteam,@Student_haclogin,@Student_hacpassword,@Student_contactid,@Student_mealstatus,@Timestamp)'
    if (@('mysql','mariadb') -contains $database.dbtype) {
      $query = 'INSERT INTO `students_extras`',$query,'ON DUPLICATE KEY UPDATE `Student_id`=@Student_id,`Student_gradyr`=@Student_gradyr,`Student_nickname`=@Student_nickname,`Student_homeroom`=@Student_homeroom,`Student_hrmtid`=@Student_hrmtid,`Student_advisor`=@Student_advisor,`Student_houseteam`=@Student_houseteam,`Student_haclogin`=@Student_haclogin,`Student_hacpassword`=@Student_hacpassword,`Student_contactid`=@Student_contactid,`Student_mealstatus`=@Student_mealstatus,`Timestamp`=@Timestamp'
      Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
    } else {
      $query = 'INSERT OR REPLACE INTO `students_extras`',$query
      Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
    }
  } catch {
    write-host "Error: Could not import students_extras.csv correctly. $PSItem" -ForegroundColor RED
    exit(1)
  }
}

Complete-SqlTransaction

if (-Not($QuickMode)) { #Quickmode should stop here.

  Start-SqlTransaction
  
  write-host "Info: Importing schools.csv."
  Import-CSV .\files\schools.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`School_id`,`School_name`,`School_number`,`State_id`,`Low_grade`,`High_grade`,`Principal`,`Principal_email`,`School_address`,`School_city`,`School_state`,`School_zip`,`School_phone`,`Timestamp`)
      VALUES
      (@School_id,@School_name,@School_number,@State_id,@Low_grade,@High_grade,@Principal,@Principal_email,@School_address,@School_city,@School_state,@School_zip,@School_phone,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `schools`',$query,'ON DUPLICATE KEY UPDATE `School_id`=@School_id,`School_name`=@School_name,`School_number`=@School_number,`State_id`=@State_id,`Low_grade`=@Low_grade,`High_grade`=@High_grade,`Principal`=@Principal,`Principal_email`=@Principal_email,`School_address`=@School_address,`School_city`=@School_city,`School_state`=@School_state,`School_zip`=@School_zip,`School_phone`=@School_phone,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `schools`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import schools.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  write-host "Info: Importing teachers.csv."
  Import-CSV .\files\teachers.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { 
      if ($_.Name -eq 'State_teacher_id') { $_.Value = [int64]$_.Value }
      $hashtable[$_.Name] = $_.Value
    }
    try {
      $query = '(`School_id`,`Teacher_id`,`Teacher_number`,`State_teacher_id`,`Teacher_email`,`First_name`,`Middle_name`,`Last_name`,`Title`,`Username`,`Password`,`Timestamp`)
      VALUES
      (@School_id,@Teacher_id,@Teacher_number,@State_teacher_id,@Teacher_email,@First_name,@Middle_name,@Last_name,@Title,@Username,@Password,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `teachers`',$query,'ON DUPLICATE KEY UPDATE `School_id`=@School_id,`Teacher_id`=@Teacher_id,`Teacher_number`=@Teacher_number,`State_teacher_id`=@State_teacher_id,`Teacher_email`=@Teacher_email,`First_name`=@First_name,`Middle_name`=@Middle_name,`Last_name`=@Last_name,`Title`=@Title,`Username`=@Username,`Password`=@Password,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `teachers`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import teachers.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  write-host "Info: Importing teachers_extras.csv."
  Import-CSV .\files\teachers_extras.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    try {
      $hashtable = @{}
      $PSItem.psobject.properties | ForEach-Object {
        if ($_.Name -eq 'State_id') { $_.Value = [int64]$_.Value }
        $hashtable[$_.Name] = $_.Value
      }
      Invoke-SqlUpdate -Query $('REPLACE INTO `teachers_extras` (' + $('`' + ($hashtable.Keys -join '`,`') + '`') + ') VALUES (' +  $('@' + ($hashtable.Keys -join ',@')) + ')') -Parameters $hashtable | Out-Null
    } catch {
      write-host "Error: Could not import teachers_extras.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  write-host "Info: Importing staff.csv."
  Import-CSV .\files\staff.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    try {
      $hashtable = @{}
      $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
      Invoke-SqlUpdate -Query $('REPLACE INTO `staff` (' + $('`' + ($hashtable.Keys -join '`,`') + '`') + ') VALUES (' +  $('@' + ($hashtable.Keys -join ',@')) + ')') -Parameters $hashtable | Out-Null
    } catch {
      write-host "Error: Could not import staff.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }


  write-host "Info: Importing hac_students.csv."
  Import-CSV .\files\hac_students.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`Contact_id`,`Student_id`,`Student_haclogin`,`Student_webaccess`,`Student__lastpasswordchange`,`Student_lastlogindate`,`Student_accesstoken`,`Student_accesstokenused`,`Timestamp`)
      VALUES
      (@Contact_id,@Student_id,@Student_haclogin,@Student_webaccess,@Student__lastpasswordchange,@Student_lastlogindate,@Student_accesstoken,@Student_accesstokenused,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `hac_students`',$query,'ON DUPLICATE KEY UPDATE `Contact_id`=@Contact_id,`Student_id`=@Student_id,`Student_haclogin`=@Student_haclogin,`Student_webaccess`=@Student_webaccess,`Student__lastpasswordchange`=@Student__lastpasswordchange,`Student_lastlogindate`=@Student_lastlogindate,`Student_accesstoken`=@Student_accesstoken,`Student_accesstokenused`=@Student_accesstokenused,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `hac_students`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import hac_students.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  #Complete-SqlTransaction
  #Start-SqlTransaction

  write-host "Info: Importing hac_guardians.csv."
  Import-CSV .\files\hac_guardians.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`Contact_id`,`Student_id`,`Guardian_firstname`,`Guardian_lastname`,`Guardian_priority`,`Guardian_haclogin`,`Guardian_webaccess`,`Guardian_passchangenextlogin`,`Guardian_lastpasswordchange`,`Guardian_lastlogindate`,`Guardian_accesstoken`,`Guardian_accesstokenused`,`Timestamp`)
      VALUES
      (@Contact_id,@Student_id,@Guardian_firstname,@Guardian_lastname,@Guardian_priority,@Guardian_haclogin,@Guardian_webaccess,@Guardian_passchangenextlogin,@Guardian_lastpasswordchange,@Guardian_lastlogindate,@Guardian_accesstoken,@Guardian_accesstokenused,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `hac_guardians`',$query,'ON DUPLICATE KEY UPDATE `Contact_id`=@Contact_id,`Student_id`=@Student_id,`Guardian_firstname`=@Guardian_firstname,`Guardian_lastname`=@Guardian_lastname,`Guardian_priority`=@Guardian_priority,`Guardian_haclogin`=@Guardian_haclogin,`Guardian_webaccess`=@Guardian_webaccess,`Guardian_passchangenextlogin`=@Guardian_passchangenextlogin,`Guardian_lastpasswordchange`=@Guardian_lastpasswordchange,`Guardian_lastlogindate`=@Guardian_lastlogindate,`Guardian_accesstoken`=@Guardian_accesstoken,`Guardian_accesstokenused`=@Guardian_accesstokenused,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `hac_guardians`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import hac_guardians.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  write-host "Info: Importing contacts.csv."
  Import-CSV .\files\contacts.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | Where-Object { $PSItem.Contact_id -ge 1 } | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`Contact_id`,`Student_id`,`Contact_priority`,`Contact_firstname`,`Contact_lastname`,`Contact_type`,`Contact_phonenumber`,`Contact_relationship`,`Contact_phonetype`,`Contact_email`,`Timestamp`)
      VALUES
      (@Contact_id,@Student_id,@Contact_priority,@Contact_firstname,@Contact_lastname,@Contact_type,@Contact_phonenumber,@Contact_relationship,@Contact_phonetype,@Contact_email,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `contacts`',$query,'ON DUPLICATE KEY UPDATE `Contact_id`=@Contact_id,`Student_id`=@Student_id,`Contact_priority`=@Contact_priority,`Contact_firstname`=@Contact_firstname,`Contact_lastname`=@Contact_lastname,`Contact_type`=@Contact_type,`Contact_phonenumber`=@Contact_phonenumber,`Contact_relationship`=@Contact_relationship,`Contact_phonetype`=@Contact_phonetype,`Contact_email`=@Contact_email,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `contacts`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import contacts.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  #Complete-SqlTransaction
  #Start-SqlTransaction

  write-host "Info: Importing transportation.csv."
  Import-CSV .\files\transportation.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`Student_id`,`Student_BusNumFrom`,`Student_TravelTypeFrom`,`Student_BusNumTo`,`Student_TravelTypeTo`,`Timestamp`)
      VALUES
      (@Student_id,@Student_BusNumFrom,@Student_TravelTypeFrom,@Student_BusNumTo,@Student_TravelTypeTo,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `transportation`',$query,'ON DUPLICATE KEY UPDATE `Student_id`=@Student_id,`Student_BusNumFrom`=@Student_BusNumFrom,`Student_TravelTypeFrom`=@Student_TravelTypeFrom,`Student_BusNumTo`=@Student_BusNumTo,`Student_TravelTypeTo`=@Student_TravelTypeTo,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `transportation`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import transportation.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  #Complete-SqlTransaction
  #Start-SqlTransaction

  write-host "Info: Importing sections.csv."
  Import-CSV .\files\sections.csv | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`School_id`,`Section_id`,`Teacher_id`,`Teacher_2_id`,`Teacher_3_id`,`Teacher_4_id`,`Name`,`Section_number`,`Grade`,`Course_name`,`Course_number`,`Course_description`,`Period`,`Subject`,`Term_name`,`Term_start`,`Term_end`,`Timestamp`)
      VALUES
      (@School_id,@Section_id,@Teacher_id,@Teacher_2_id,@Teacher_3_id,@Teacher_4_id,@Name,@Section_number,@Grade,@Course_name,@Course_number,@Course_description,@Period,@Subject,@Term_name,@Term_start,@Term_end,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `sections`',$query,'ON DUPLICATE KEY UPDATE `School_id`=@School_id,`Section_id`=@Section_id,`Teacher_id`=@Teacher_id,`Teacher_2_id`=@Teacher_2_id,`Teacher_3_id`=@Teacher_3_id,`Teacher_4_id`=@Teacher_4_id,`Name`=@Name,`Section_number`=@Section_number,`Grade`=@Grade,`Course_name`=@Course_name,`Course_number`=@Course_number,`Course_description`=@Course_description,`Period`=@Period,`Subject`=@Subject,`Term_name`=@Term_name,`Term_start`=@Term_start,`Term_end`=@Term_end,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `sections`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import sections.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  #Complete-SqlTransaction

  #Importing Enrollments takes way to long. This file can have hundreds of thousands of records.
  #We need to try and import the CSV file to a temporary table then bulkcopy to the database.
  #I have tried to import to :memory: but it takes 45x longer than importing the CSV directly.
  write-host "Info: Importing enrollments.csv."
  #first try doing our CSV import. Then revert to manually importing one at a time.

  try {

    $enrollments_csv_import = @"
drop table if exists enrollments_csv_import;
.mode csv
.separator ,
.import files\\enrollments.csv enrollments_csv_import
"@

  if (@('mysql','mariadb') -contains $database.dbtype) {
      #load the csv into a temporary sqlite3 database so we can bulk copy to our database.
      $enrollments_csv_import | & $currentPath\bin\sqlite3.exe $currentPath\enrollments.sqlite3
      Open-SQLiteConnection -ConnectionName "enrollments_temp_db" -DataSource .\enrollments.sqlite3
      Invoke-SqlUpdate -Query 'DROP TABLE IF EXISTS `enrollments_csv_import`;' | Out-Null
      Invoke-SqlUpdate -Query 'CREATE TABLE `enrollments_csv_import` (
        `School_id` int(4),
        `Section_id` int(32),
        `Student_id` int(10),
        `Marking_period` int(1)
      );' | Out-Null
      Invoke-SqlBulkCopy -SourceConnectionName "enrollments_temp_db" -SourceTable "enrollments_csv_import" -DestinationConnectionName "default" -DestinationTable "enrollments_csv_import" | Out-Null
      #Close-SqlConnection -ConnectionName "enrollments_temp_db"

      #Start-SqlTransaction
      Invoke-SqlUpdate -Query "INSERT INTO ``enrollments`` SELECT *,$Timestamp as ``Timestamp`` FROM ``enrollments_csv_import`` ON DUPLICATE KEY UPDATE ``Timestamp``=$Timestamp" | Out-Null
      #Complete-SqlTransaction

    } elseif ((Get-SqlConnection).FileName) {
      #sqlite into the same database.
      Complete-SqlTransaction
      $enrollments_csv_import | & $currentPath\bin\sqlite3.exe (Get-SqlConnection).FileName
      Start-SqlTransaction
      Invoke-SqlUpdate -Query "INSERT OR REPLACE INTO enrollments SELECT *,$Timestamp AS Timestamp FROM enrollments_csv_import;" | Out-Null
    }
  } catch {
    write-host "Error: Could not import enrollments.csv correctly. $PSItem" -ForegroundColor RED
    exit(1)
  }

  #Start-SqlTransaction

  write-host "Info: Importing activities.csv."
  $activities = Import-CSV .\files\activities.csv | Where-Object { $activitiesBuildings -contains $PSItem.'School_id' }
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
  $activities | Select-Object -Property *,@{Name='Timestamp';Expression={ $timestamp }} | ForEach-Object {
    $hashtable = @{}
    $PSItem.psobject.properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }
    try {
      $query = '(`School_id`,`Student_id`,`Teacher_id`,`Activity_code`,`Activity_name`,`Timestamp`)
      VALUES
      (@School_id,@Student_id,@Teacher_id,@Activity_code,@Activity_name,@Timestamp)'
      if (@('mysql','mariadb') -contains $database.dbtype) {
        $query = 'INSERT INTO `activities`',$query,'ON DUPLICATE KEY UPDATE `School_id`=@School_id,`Student_id`=@Student_id,`Teacher_id`=@Teacher_id,`Activity_code`=@Activity_code,`Activity_name`=@Activity_name,`Timestamp`=@Timestamp'
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      } else {
        $query = 'INSERT OR REPLACE INTO `activities`',$query
        Invoke-SqlUpdate -Query $query -Parameters $hashtable | Out-Null
      }
    } catch {
      write-host "Error: Could not import activities.csv correctly. $PSItem" -ForegroundColor RED
      exit(1)
    }
  }

  Complete-SqlTransaction
  Start-SqlTransaction

  ###########################################################
  # Pull Term Data from imported Sections and select currentTerm.
  ###########################################################
  $term1 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 1 LIMIT 1"
  $term2 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 2 LIMIT 1"
  $term3 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 3 LIMIT 1"
  $term4 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections WHERE Term_name = 4 LIMIT 1"

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

  if ($Term) {
      $currentTerm = $Term
  }

  Complete-SqlTransaction
}


Start-SqlTransaction

###########################################################
# Remove old teachers, students, sections, enrollments
###########################################################
if (-Not($SkipTableCleanup)) {
  write-host "Info: Archiving and deleting old entries. DOES NOT INCLUDE students_extras OR activities tables!"
  #Archive students not in CSV
  if ($RetainStudentsDays -ge 1) {
    Invoke-SqlUpdate -Query "INSERT INTO students_disabled SELECT * FROM students WHERE Timestamp < $timestampRetainStudents; `
    DELETE FROM students WHERE Timestamp < $timestampRetainStudents;" | Out-Null
  } else {
    $q = 'INSERT INTO students_disabled SELECT * FROM students WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1);
    DELETE FROM students WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1) AS X);'
    Invoke-SqlUpdate -Query $q | Out-Null
  }

    if (-Not($QuickMode)) {
    #delete teachers no longer in CSV
    Invoke-SqlUpdate -Query 'DELETE FROM teachers WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM teachers ORDER BY Timestamp DESC LIMIT 1) AS X);' | Out-Null

    #delete contacts no longer in CSV
    Invoke-SqlUpdate -Query 'DELETE FROM contacts WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM contacts ORDER BY Timestamp DESC LIMIT 1) AS X);' | Out-Null

    #delete transportation no longer in CSV
    Invoke-SqlUpdate -Query 'DELETE FROM transportation WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM transportation ORDER BY Timestamp DESC LIMIT 1) AS X);' | Out-Null

    #remove invalid buildings that might pull in with the staff.
    Invoke-SqlUpdate -Query 'DELETE FROM teachers WHERE School_id NOT IN (SELECT School_id from schools)' | Out-Null
    Invoke-SqlUpdate -Query 'DELETE FROM staff WHERE School_id NOT IN (SELECT School_id from schools)' | Out-Null
    
    #remove duplicate staff that are also listed in teachers.
    Invoke-SqlUpdate -Query 'DELETE FROM `staff` WHERE Staff_id IN (SELECT `Teacher_id` FROM `teachers`)' | Out-Null

    #delete sections no longer in CSV
    Invoke-SqlUpdate -Query 'DELETE FROM sections WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM sections ORDER BY Timestamp DESC LIMIT 1) AS X);' | Out-Null

    #archive enrollments no longer in CSV
    if (@('mysql','mariadb') -contains $database.dbtype) {
      Invoke-SqlUpdate -Query "INSERT IGNORE INTO enrollments_archived SELECT * FROM enrollments WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1);" | Out-Null
    } else {
      $query = 'INSERT OR REPLACE INTO `students`',$query
      Invoke-SqlUpdate -Query "INSERT OR REPLACE INTO enrollments_archived SELECT * FROM enrollments WHERE Timestamp != (SELECT DISTINCT Timestamp FROM students ORDER BY Timestamp DESC LIMIT 1);" | Out-Null
    }

    #delete hac_guardians no longer in CSV
    Invoke-SqlUpdate -Query 'DELETE FROM hac_guardians WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM hac_guardians ORDER BY Timestamp DESC LIMIT 1) AS X);' | Out-Null

    #delete hac_students no longer in CSV
    Invoke-SqlUpdate -Query 'DELETE FROM hac_students WHERE Student_id NOT IN (SELECT DISTINCT Student_id FROM students);' | Out-Null

    #archive, retrieve, retain for X days and on only certain buildings. 
    if ($RetainEnrollmentsDays -ge 1) {

      #Just in case this setting is changed we need to pull back in from the enrollments_archived table.
      Invoke-SqlUpdate -Query "INSERT INGORE enrollments SELECT * FROM enrollments_archived WHERE Timestamp > $timestampRetain;" | Out-Null

      #Clean up older enrollments from enrollments table.
      Invoke-SqlUpdate -Query "DELETE FROM enrollments WHERE Timestamp < $timestampRetain;" | Out-Null

      #We need to delete older enrollments for buildings not specified in $RetainEnrollmentBuildings
      if ($RetainEnrollmentBuildings) {
        $RetainBuildingsString = "`($($($RetainEnrollmentBuildings) -join (','))`)"
        Invoke-SqlUpdate -Query "DELETE FROM enrollments WHERE Timestamp != (SELECT DISTINCT Timestamp FROM enrollments ORDER BY Timestamp DESC LIMIT 1) AND School_id NOT IN $RetainBuildingsString;" | Out-Null
      }

    } else {
      #Delete everything older.
      Invoke-SqlUpdate -Query 'DELETE FROM enrollments WHERE Timestamp != (SELECT * FROM (SELECT DISTINCT Timestamp FROM enrollments ORDER BY Timestamp DESC LIMIT 1) AS X);' | Out-Null
    }

    #Deleting archived information that is older than 1 year. This is really to save yourself from yourself if you set a ridiculous $timestampRetain value.
    Invoke-SqlUpdate -Query "DELETE FROM enrollments WHERE Timestamp < $timestamp365DaysAgo;" | Out-Null
    Invoke-SqlUpdate -Query "DELETE FROM enrollments_archived WHERE Timestamp < $timestamp365DaysAgo;" | Out-Null

    #Mark Passwords of students that are inactive to 2 so we don't query them later.
    Invoke-SqlUpdate -Query "UPDATE passwords SET HAC_passwordset = 2 WHERE Student_id NOT IN (SELECT Student_id FROM students)" | Out-Null
    #General passwords cleanup of inactive students passwords.
    Invoke-SqlUpdate -Query "DELETE FROM passwords WHERE Timestamp < $timestamp365DaysAgo AND HAC_passwordset = 2;" | Out-Null

  }
}

Complete-SqlTransaction

######################################################################################################################
# Group Enrollments and Sections together to match for an actual Schedule.
######################################################################################################################
if (-Not($QuickMode)) {
  write-host "Info: Building Merged Sections and Enrollments to build correct Schedules."

  #build schedules based on matching enrollments
  try {
    Invoke-SqlUpdate -Query '/* FINAL REVISION! */
    /* To ensure the table columns are correct I am dropping the entire table instead of truncating. */
    DROP TABLE IF EXISTS `enrollments_grouped`;
    CREATE TABLE `enrollments_grouped` (
      `School_id`	int(4),
      `Section_id` bigint(32),
      `Student_id` int(10),
      `Terms`	int(4),
      UNIQUE (`School_id`,`Section_id`,`Student_id`,`Terms`)
    );' | Out-Null

    Start-SqlTransaction
    if (@('mysql','mariadb') -contains $database.dbtype) {
      Invoke-SqlUpdate -Query "SET SESSION sql_mode=(SELECT REPLACE(@@sql_mode,'ONLY_FULL_GROUP_BY',''));" | Out-Null
      $group_concat = "group_concat(Marking_period ORDER BY Marking_period SEPARATOR '')"
    } else {
      $group_concat = "group_concat(Marking_period,'')"
    }
    Invoke-SqlUpdate -Query '/* I need the terms to be grouped together to query later. Make a copy of the enrollments table with all terms grouped together. */
    REPLACE INTO `enrollments_grouped`
      SELECT enrollments.School_id,Section_id,enrollments.Student_id,',$group_concat,' as Terms
      FROM enrollments
      LEFT JOIN students ON enrollments.Student_id = students.Student_id
      GROUP BY enrollments.Student_id,enrollments.Section_id
      ORDER BY Marking_period;' | Out-Null
    Complete-SqlTransaction

    Invoke-SqlUpdate -Query '/* I need the terms to be grouped together to query later. Make a copy of the sections table with all terms grouped together. */
    DROP TABLE IF EXISTS `sections_grouped`;
    CREATE TABLE `sections_grouped` (
      `Terms` int(4),
      `School_id` int(4),
      `Section_id` bigint(32),
      `Teacher_id` varchar(32),
      `Name` varchar(64),
      `Section_number` varchar(32),
      `Grade` varchar(15),
      `Course_name` varchar(64),
      `Course_number` varchar(32),
      `Course_description` varchar(64),
      `Period` varchar(10),
      `Subject` varchar(32),
      `Term_start` varchar(10),
      `Term_end` varchar(10),
      UNIQUE (`Section_id`,`Terms`)
    );' | Out-Null

    Start-SqlTransaction
    if (@('mysql','mariadb') -contains $database.dbtype) {
      Invoke-SqlUpdate -Query "SET SESSION sql_mode=(SELECT REPLACE(@@sql_mode,'ONLY_FULL_GROUP_BY',''));" | Out-Null
      $group_concat = "group_concat(Term_name ORDER BY Term_name SEPARATOR '')"
    } else {
      $group_concat = "group_concat(Term_name,'')"
    }
    Invoke-SqlUpdate -Query 'REPLACE INTO sections_grouped
      SELECT ',$group_concat,' as `Terms`,`School_id`, `Section_id`, `Teacher_id`, `Name`, `Section_number`, `Grade`, `Course_name`, `Course_number`, `Course_description`, `Period`, `Subject`, `Term_start`, `Term_end`
      FROM `sections`
      GROUP BY `sections`.`Section_id`
      ORDER BY `Terms`' | Out-Null
    Complete-SqlTransaction

    # This table will hold our final results. Matching all kids who are enrolled for the entire length of the section terms. Leaving out classes that do not have a teacher email assigned to them.
    Invoke-SqlUpdate -Query 'DROP TABLE IF EXISTS `schedules`;
    CREATE TABLE `schedules` (
      `School_id` int(4),
      `Section_id` bigint(32),
      `Course_number` varchar(32),
      `Section_number` bigint(12),
      `Terms` int(4),
      `Student_id` int(10),
      `Teacher_id` varchar(32),
      `Period` varchar(12),
      `First_name` varchar(64),
      `Last_name` varchar(64),
      `Student_email` varchar(128),
      `Teacher_firstname` varchar(64),
      `Teacher_lastname` varchar(64),
      `Teacher_email` varchar(128),
      UNIQUE (`School_id`,`Section_id`,`Student_id`)
    );' | Out-Null

    Start-SqlTransaction
    Invoke-SqlUpdate -Query '/* This requires that teachers have an email address. This is on purpose for my district because of sections that are placeholders. */
      INSERT INTO `schedules`
      SELECT `enrollments_grouped`.`School_id`, `enrollments_grouped`.`Section_id`, `sections_grouped`.`Course_number`, `sections_grouped`.`Section_number`, `enrollments_grouped`.`Terms`, `enrollments_grouped`.`Student_id`, `sections_grouped`.`Teacher_id`, `sections_grouped`.`Period`, `students`.`First_name`, `students`.`Last_name`, `students`.`Student_email`, `teachers`.`First_name` as `Teacher_firstname`, `teachers`.`Last_name` as `Teacher_lastname`, `teachers`.`Teacher_email`
      FROM `enrollments_grouped`
      INNER JOIN `sections_grouped` ON `enrollments_grouped`.`Section_id` = `sections_grouped`.`Section_id` AND `enrollments_grouped`.`Terms` = `sections_grouped`.`Terms`
      INNER JOIN `students` ON `students`.`Student_id` = `enrollments_grouped`.`Student_id`
      INNER JOIN `teachers` ON `sections_grouped`.`Teacher_id` = `teachers`.`Teacher_id` AND `enrollments_grouped`.`School_id` = `teachers`.`School_id`
      WHERE `Teacher_email` != '''' ORDER BY `enrollments_grouped`.`Student_id`,`sections_grouped`.`Period`;' | Out-Null
    Complete-SqlTransaction

  } catch {
    write-host "ERROR: Failed to create merged sections and enrollments! $_" -ForegroundColor RED
    exit(1)
  }

  Start-SqlTransaction

  if (@('mysql','mariadb') -contains $database.dbtype) {
    Invoke-SqlUpdate -Query "SET SESSION sql_mode=(SELECT REPLACE(@@sql_mode,'ONLY_FULL_GROUP_BY',''));" | Out-Null
  }

  #Now insert just the current Term to fix enrollment. We ignore duplicates
  if (@('mysql','mariadb') -contains $database.dbtype) {
    $q = 'REPLACE INTO `schedules`
    SELECT enrollments.School_id,enrollments.Section_id, sections.Course_number, sections.Section_number, enrollments.Marking_period as `Terms`, enrollments.Student_id, sections.Teacher_id, sections.Period, students.First_name, students.Last_name, students.Student_email, teachers.First_name as Teacher_firstname, teachers.Last_name as Teacher_lastname, teachers.Teacher_email
    FROM enrollments
    INNER JOIN sections ON enrollments.Section_id = sections.Section_id
    INNER JOIN students ON students.Student_id = enrollments.Student_id
    INNER JOIN teachers ON sections.Teacher_id = teachers.Teacher_id AND enrollments.School_id = teachers.School_id
    WHERE Teacher_email != '''' AND Marking_period =',$currentTerm,' ORDER BY enrollments.Student_id,sections.Period'
  } else {
    $q = "INSERT OR REPLACE INTO schedules SELECT enrollments.School_id,enrollments.Section_id, sections.Course_number, sections.Section_number, enrollments.Marking_period as 'Terms', enrollments.Student_id, sections.Teacher_id, sections.Period, students.First_name, students.Last_name, students.Student_email, teachers.First_name as Teacher_firstname, teachers.Last_name as Teacher_lastname, teachers.Teacher_email FROM enrollments INNER JOIN sections ON enrollments.Section_id = sections.Section_id INNER JOIN students ON students.Student_id = enrollments.Student_id INNER JOIN teachers ON sections.Teacher_id = teachers.Teacher_id AND enrollments.School_id = teachers.School_id WHERE Teacher_email != '' AND Marking_period = $currentTerm ORDER BY enrollments.Student_id,sections.Period"
  }

  try {
    Invoke-SqlUpdate -Query $q  | Out-Null
  } catch {
    write-host "ERROR: Failed to create schedules based on merged sections and enrollments!" -ForegroundColor RED
    exit(1)
  }

  Complete-SqlTransaction

} #Close QuickMode

#Verify we have content in the grouped tables and schedules. We have to know we have data in there.
write-host "Info: Verifying tables meet the minimum students count of $minStuCount specified in settings.ps1"
try {
  if ((Invoke-SqlScalar -Query "SELECT COUNT(*) AS count FROM students") -le $minStuCount) { write-host "ERROR: Not enough students to pass verification. Please double check your import file for a students." -ForegroundColor RED; exit(1) }
  if ((Invoke-SqlScalar -Query "SELECT COUNT(*) AS count FROM sections_grouped") -le $minStuCount) { write-host "ERROR: Not enough sections_grouped to pass verification. Please double check your import file for a sections." -ForegroundColor RED; exit(1) }
  if ((Invoke-SqlScalar -Query "SELECT COUNT(*) AS count FROM enrollments_grouped") -le $minStuCount) { write-host "ERROR: Not enough enrollments_grouped to pass verification. Please double check your import file for a enrollments." -ForegroundColor RED; exit(1) }
  if ((Invoke-SqlScalar -Query "SELECT COUNT(*) AS count FROM schedules") -le $minStuCount) { write-host "ERROR: Not enough schedules to pass verification. Please double check your import file for a schedules." -ForegroundColor RED; exit(1) }
} catch {
  write-host "ERROR: Failed to verify count in students, grouped sections, grouped enrollments, and/or schedules." -ForegroundColor RED
  exit(1)
}

Close-SqlConnection

# Final Script to do all the stuff.
if (-Not($DisablePostProcessingScript)) {
  if (-Not($Staging)) {
      if (Test-Path $currentPath\x_PostProcessingDatabase.ps1) {
        if ($QuickMode) {
          . $currentPath\x_PostProcessingDatabase.ps1 -QuickMode #I still want it in this code block after checking for x_PostProcessingDatabase.ps1. QuickMode should only run on a fully configured system.
        } else {
          . $currentPath\x_PostProcessingDatabase.ps1
        }
      }
  } else {
      Write-Host "Info: Staging specified. Will not run x_PostProcessingDatabase.ps1."
  }
} else {
  Write-Host "Info: DisablePostProcessingScript specified. Will not run x_PostProcessingDatabase.ps1."
}

Stop-TranScript