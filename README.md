Automated_Students
==================

This is the repository for the public version of the Automated_Students project. It does not contain any additional scripts besides the required ones to run the project. If you're interested in having your district automated with a lot of additional functionality please contact Craig Millsap.

This project started as an student account automation project using just the basic Clever files. However, that turned out to be insufficient.
So a few modified Cognos reports and queries and this is the final project. The reports are located in 
> Cognos > Team Content > Student Management System > _Shared Data File Reports > /automated_students > [version]

There are two versions of this program.
One is licensed GPL and in the AR-k12code repository.
The other is a paid service of CAMTech Computer Serices, LLC (Craig Millsap).

Student Accounts are matched on their school assigned Student ID to the EmployeeNumber field in AD.

# Scripts Features
- Automatic Active Directory Structure Creation. (choose from 3)
- Automatic UPN Suffix creation needed for M365. (choose per building or for all students)
- Student account creation (you choose the design of the username from 6 options)
- Notifications of new student accounts.
- Generated passwords saved to a CSV (Sample Code to share with Scheduled Teachers via Google Drive).
- Copying data to Google Drive and updating Spreadsheets with new custom CSV data.
- Student Management Groups and permissions on AD. (Where staff can reset passwords for students.)
- Distribution Groups for Students based on building and grade level.
- Managing owners of the distribution groups and syncing to Google Workspaces. (Including class sponsors for certain grades)
- Automatic name change detection and name conflict resolution. 
- Automatically move students between building OU and GradYR.
- Managed Home Directories and moves them when a student changes buildings if needed.
- Staging to see what will be created if the script is run in production. No changes will be done to the Domain (Please be careful with the custom scripts for your school.)
- Parameters to tell the scripts what tasks you want to run. For example, "./automated_students.ps1 -SkipStudents -DisablePostProcessingScript" or "./automated_database.ps1 -$SkipDownloadingReports"
- Examples of exporting CSV files from your own local database for importing into 3rd party programs. (Updated every time it runs.)
- Automatic export of your AD structure, usernames, objectGUIDs, etc. Just in case you need to put everything back.
- Options to build Homerooms and Activities as part of a students Schedule that is uploaded into Clever. (This helps with school startup for Elementary when schedules aren't done yet.)
- Possible integration with other projects such as the eSchoolUpload and HAC Passwords. (already built in the paid version)

Gotchas
=======
- This may restructure your AD Organizational Units if you choose something different than you already have. You should be syncing OU's to Google. Be sure to check any service that syncs or is applied to individual OU's. Think Go Guardian, Bark, Impero, etc.
- Students are matched on their Student ID to EmployeeNumber in AD. Accounts not found in eSchool will be disabled since that means they dropped from your district. If you wish to keep them active you can add those accounts to the exclusions.csv or skip disabling accounts.
- If GivenName and SurName match then a mismatch on email or username will not be evaluated. If you don't want an existing environment to be fully managed and changed then update all GivenNames and Surnames in AD with what you get from eSchool. The other fields such as email won't change.
- What is in eSchool is what you get. If bad data goes into eSchool then you're going to get bad data.
- Hyphens are a legal character in a lastname. Spaces have to be stripped automatically to build usernames. There are options in the settings.ps1 to account for this issue. Honestly though you should have some standards in eSchool.

# Requirements
Dedicated VM
 * Windows 10 or Windows Server 2019 Datacenter
 * Quad Core
 * 8GB RAM minimum.

PowerShell 7 (latest stable)
 * Download from: https://github.com/PowerShell/powershell/releases

Git for Windows
 * https://git-scm.com/download/win
 * We will use this to pull the project and any updates from github.

You will want GAM at c:\scripts\gam\gam.exe
 * Download from: https://github.com/taers232c/GAMADV-XTD3/releases
 * This is for user creation, group moderation settings, google drive uploads of CSVs.
 * User Environment variables
   1. GAMUSERCONFIGDIR = c:\scripts\gam
   2. GAMCFGDIR = c:\scripts\gam
 * Adjust path to include c:\scripts\gam
 * REBOOT AFTER ADJUSTING PATHS!!!
 * gam-setup.bat APPROVE ALL SCOPES!
 * gam user YOUR_SUPER_USER_ACCOUNT check serviceaccount (follow link and authorize domain wide delegation)

You may want rclone at c:\scripts\rclone\rclone.exe
 * Download from: https://rclone.org/downloads/
 * This is for listing google drive as json and syncing folders.
 * Adjust path to include c:\scripts\rclone
 * rclone --config c:\scripts\rclone\rclone.conf config

 You will want the SQLite DB Browser
 * https://sqlitebrowser.org/dl

Install
=======
 * Open Powershell 7 as Administrator
```
cd \
mkdir scripts
cd scripts
git clone https://github.com/AR-k12code/automated_students.git
Copy-Item C:\Scripts\automated_students\config_samples\sample_settings.ps1 c:\scripts\automated_students\settings.ps1
notepad c:\scripts\automated_students\settings.ps1
Copy-Item C:\Scripts\automated_students\scripts_samples\x_PostProcessingDatabase.ps1 C:\Scripts\automated_students\x_PostProcessingDatabase.ps1
notepad c:\scripts\automated_students\x_PostProcessingDatabase.ps1
```
 * Be sure to run the script with $Staging = $True in settings.ps1 until you're certain you're ready.


# Scripts
## settings.ps1
* Configuration file for your district. A sample file is provided called config_samples\settings_sample.ps1.
* Each setting should have some form of documentation with it.

## automated_database.ps1
* Main script to download and import all CSVs into a SQLite database. This makes queries local and custom for file exports. Run with Task Scheduler as a Domain Admin.
* At the end of execution calls x_PostProcessingDatabase.ps1. (customized ps1 script but this should call automated_students.ps1. See Example in scripts_examples folder.)

## automated_students.ps1
* This script queries the SQLite database for automating managing student Active Directory accounts.
* New student calls x_InterimProcessingNewAccounts.ps1 (customized ps1 script for you to do as you please.)
* Existing students calls x_InterimProcessingExistingAccounts.ps1 (customized ps1 script for you to do as you please.)
* At the end of execution calls x_PostProcessingAutomatedStudents.ps1 (customized ps1 script for you to do as you please.)

## exclusions.csv
> This is to exclude specific student accounts from creation, being disabled, modification (name change), or OU move.
````
Student_id,First_name,Last_name
123456, John, Doe
````

## overrides.csv
> This file is used to over ride first names, last names, Middle Initial, School_id, Grade
````
Student_id,First_name,Last_name,Middle_Initial,School_id,Grade
123456, John W, Doe, W
````

# Sample Scripts and Config
The folder "config_samples" contains samples for  settings.ps1 and Google Cloud Directory Sync.

The folder "scripts_samples" contains some sample code that you may want to do for Post Processing SQL, Post Processing Accounts, Interim Processing, etc.

# Post Processing Tasks

Post processing tasks can be anything you want or need custom to your school.

* Upload CSV files into Google Drive using gam or entire folders using rclone.
* Upload CSV contents into an existing Google Spreadsheet (This I use to do vlookups from other spreadsheets.)
* Upload Data into Shared Google Drives.
* Pull the generated email addresses and upload them back into eSchool.
* Run eSchool tasks such as generating HAC login information for students.

# PostProcessingScripts Folder

Any powershell scripts in this folder will automatically be ran after automated_students.ps1 has completed. These files are processed asynchronously alphanumerically excluding last_*.ps1.  Then it will run last_*.ps1 scripts.

# Troubleshooting

This script must be run with a dedicated Domain Admin login. Please ensure any cloud services has 2FA enabled.

Q. How do you renew your mail password?

A. Open powershell and "cd c:\scripts\automated_students"
   ".\automated_students.ps1 -RenewEmailPassword"

Q. How do I manually run the scripts?

A. Open Powershell and "cd c:\scripts\automated_students"
   ".\automated_database.ps1" or ".\automated_students.ps1"
   You can view the command parameters by typing " -" then pressing CTRL+SPACE after the script name.

Q. How can I see my database?

A. You can verify the database integrity by opening DB Browser and opening the c:\scripts\automated_students\YOURSCHOOLDSN.sqlite3 file.
   This will also help you to write your own custom SQL for custom CSV exports.

Q. I need to update to the latest version

A. Open Powershell and "cd \scripts\automated_students" then "git pull"
