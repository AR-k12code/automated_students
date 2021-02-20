### This repository has been archived. It can be used for a template for accomplishing your own automation. Due to accidentally including credentials while sharing code I have moved to a private repository. If you would like readaccess to the updated repository you may request it by contacting Craig Millsap and requesting a Read Only SSH Key.

Automated_Students
==================

This project started as an student account automation project using just the basic Clever files. However, that turned out to be insufficient.
So a few modified Cognos reports and queries and this is the final project. The reports are located in 
Cognos > Public Folders > _Share Temporarily Between Districts > Gentry > automation. (Copy this entire folder into your "My Folders")

Student Accounts are matched on their school assigned Student ID to the EmployeeNumber field in AD.

Scripts Features
================
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
- Ignore students in ou=Restricted,ou=SCHOOL,ou=Students,dc=domain,dc=com for special policies but move them between buildings.
- Managed Home Directories and moves them when a student changes buildings if needed.
- Staging to see what will be created if the script is run in production. No changes will be done to the Domain (Please be careful with the custom scripts for your school.)
- Parameters to tell the scripts what tasks you want to run. For example, "./automated_students.ps1 -SkipStudents -DisablePostProcessingScript" or "./automated_sqlite.ps1 -$SkipDownloadingReports"
- Examples of exporting CSV files from your own local database for importing into 3rd party programs. (Updated every time it runs.)
- Automatic export of your AD structure, usernames, objectGUIDs, etc. Just in case you need to put everything back.
- **CORRECT** Clever Exports. Term Enrollments are matched to Sections terms. So you end up with a full schedule.
- Options to build Homerooms and Activities as part of a students Schedule that is uploaded into Clever. (This helps with school startup for Elementary when schedules aren't done yet.)
- Check remote server in case there is a known issue in the state that we need to disable all scripts. Think upcoming eSchool and Cognos updates. (COMPLETELY OPTIONAL)

Gotchas
=======
- This may restructure your AD Organizational Units if you choose something different than you already have. You should be syncing OU's to Google. Be sure to check any service that syncs or is applied to individual OU's. Think Go Guardian, Bark, Impero, etc.
- Students are matched on their Student ID to EmployeeNumber in AD. Accounts not found in eSchool will be disabled since that means they dropped from your district. If you wish to keep them active you can add those accounts to the exclusions.csv or skip disabling accounts.
- If GivenName and SurName match then a mismatch on email or username will not be evaluated. If you don't want an existing environment to be fully managed and changed then update all GivenNames and Surnames in AD with what you get from eSchool. The other fields such as email won't change.
- What is in eSchool is what you get. If bad data goes into eSchool then you're going to get bad data.
- Hyphens are a legal character in a lastname. Spaces have to be stripped automatically to build usernames. There are options in the settings.ps1 to account for this issue. Honestly though you should have some standards in eSchool.

Requirements
=====================================================
Dedicated VM
 * Windows 10 or Windows Server 2019 Datacenter
 * Quad Core
 * 8GB RAM minimum.

PowerShell 7 (latest stable)
 * Download from: https://github.com/PowerShell/powershell/releases

Git for Windows
 * https://git-scm.com/download/win
 * We will use this to pull the project and any updates from github.

You need GAM at c:\scripts\gam\gam.exe
 * Download from: https://github.com/taers232c/GAMADV-XTD3/releases
 * This is for user creation, group moderation settings, google drive uploads of CSVs.
 * User Environment variables
   1. GAMUSERCONFIGDIR = c:\scripts\gam
   2. GAMCFGDIR = c:\scripts\gam
 * Adjust path to include c:\scripts\gam
 * gam create project
 * gam oauth create
    1. [R]  0)  Calendar API (supports readonly)
    2. [*]  1)  Classroom API - Course Announcements (supports readonly)
    3. [*]  2)  Classroom API - Course Topics (supports readonly)
    4. [*]  3)  Classroom API - Course Work/Submissions (supports readonly)
    5. [*]  4)  Classroom API - Courses (supports readonly)
    6. [*]  5)  Classroom API - Profile Emails
    7. [*]  6)  Classroom API - Profile Photos
    8. [*]  7)  Classroom API - Rosters (supports readonly)
    9. [*]  8)  Classroom API - Student Guardians (supports readonly)
    10. [*] 14)  Directory API - Chrome OS Devices (supports readonly)
    11. [R] 15)  Directory API - Customers (supports readonly)
    12. [R] 16)  Directory API - Domains (supports readonly)
    13. [*] 17)  Directory API - Groups (supports readonly)
    14. [R] 18)  Directory API - Mobile Devices Directory (supports readonly and action)
    15. [*] 19)  Directory API - Organizational Units (supports readonly)
    16. [R] 21)  Directory API - Roles (supports readonly)
    17. [R] 22)  Directory API - User Schemas (supports readonly)
    18. [*] 23)  Directory API - User Security
    19. [*] 24)  Directory API - Users (supports readonly)
    20. [*] 27)  Groups Settings API
    21. [*] 32)  Reports API - Usage Reports
 * gam user YOUR_SUPER_USER_ACCOUNT check serviceaccount (follow link and authorize domain wide delegation)

You need rclone at c:\scripts\rclone\rclone.exe
 * Download from: https://rclone.org/downloads/
 * This is for listing google drive as json and syncing folders.
 * Adjust path to include c:\scripts\rclone
 * rclone --config c:\scripts\rclone\rclone.conf config

 You will want the SQLite DB Browser
 * https://sqlitebrowser.org/dl

Install
=======
 * Open Powershell 7 as Administrator
```cd \
mkdir scripts
cd scripts
git clone https://github.com/AR-k12code/automated_students.git
Copy-Item C:\Scripts\automated_students\config_samples\sample_settings.ps1 c:\scripts\automated_students\settings.ps1
notepad c:\scripts\automated_students\settings.ps1
Copy-Item C:\Scripts\automated_students\scripts_samples\x_PostProcessingSQLite.ps1 C:\Scripts\automated_students\x_PostProcessingSQLite.ps1
notepad x_PostProcessingSQLite.ps1
```
 * Be sure to run the script with $Staging set until you're certain you're ready. This is in the settings.ps1 file as well towards the bottom.


Scripts
=======
settings.ps1 - Configuration file for your district. A sample file is provided called config_samples\settings_sample.ps1.
 * Each setting should have some form of **limited** documentation with it.

automated_sqlite.ps1 - Main script to import all CSVs into a SQLite database. This makes queries local and custom for file exports. Run with Task Scheduler as a Domain Admin.
 * At the end of execution calls x_PostProcessingSQLite.ps1. (customized ps1 script but this should call automated_students.ps1. See Example.)

automated_students.ps1 - This script queries the SQLite database for automating managing student Active Directory accounts.
 * New student calls x_InterimProcessingNewAccounts.ps1 (customized ps1 script for you to do as you please.)
 * Existing students calls x_InterimProcessingExistingAccounts.ps1 (customized ps1 script for you to do as you please.)
 * At the end of execution calls x_PostProcessingAutomatedStudents.ps1 (customized ps1 script for you to do as you please.)

clever.ps1 - This script pulls from the SQLite database the files required for Clever. This also fixes the enrollment problem with switching between terms.
 * This should call x_InterimProcessingClever.ps1 (A custom powershell script. I use this to transfer the Clever files to ScriptK12 as well.)

exclusions.csv - This is to exclude specific student accounts from creation, being disabled, modification (name change), or OU move.
 * Headers: Student_id,First_name,Last_name

Sample Scripts and Config
=========================
The folder "config_samples" contains samples for  settings.ps1 and Google Cloud Directory Sync.

The folder "scripts_samples" contains some sample code that you may want to do for Post Processing SQL, Post Processing Accounts, Interim Processing, etc.

Post Processing Tasks
=====================
Post processing tasks can be anything you want or need custom to your school.

* Upload CSV files into Google Drive using gam or entire folders using rclone.
* Upload CSV contents into an existing Google Spreadsheet (This I use to do vlookups from other spreadsheets.)
* Upload Data into Shared Google Drives.
* Pull the generated email addresses and upload them back into eSchool.
* Run eSchool tasks such as generating HAC login information for students.

To Do
=====
- [ ] Migrate from using pscp.exe to Posh-SSH

Troubleshooting
===============
This script must be run with a dedicated Domain Admin login. Please 2 factor cloud services.

Q. How do you renew your Cognos password?

A. Delete the c:\scripts\apscnpw.txt file.
   Open powershell and "cd c:\scripts"
   ".\CognosDownload.ps1 -username 0000username -report xyz"
   Save your new password when prompted. Report will error but you can ignore.

Q. How do you renew your mail password?

A. Open powershell and "cd c:\scripts\automated_students"
   ".\automated_students.ps1 -RenewEmailPassword"

Q. How do I manually run the scripts?

A. Open Powershell and "cd c:\scripts\automated_students"
   ".\automated_sqlite.ps1" or ".\automated_students.ps1"
   You can view the command parameters by typing " -" then pressing CTRL+SPACE after the script name.

Q. How can I see my database?

A. You can verify the database integrity by opening DB Browser and opening the c:\scripts\automated_students\YOURSCHOOLDSN.sqlite3 file.
   This will also help you to write your own custom SQL for custom CSV exports.

Q. Trick the system into reevaluating home directories and permissions?

A. If a name change is detected it will force checking all home directory permissions. Adjust for your school but the following would change all names on accounts. Running the script again would fix it. WARNING: This might also change email addresses if GivenName and Surname doesn't generate the same email/userprincipal name!

```Get-ADUser -Filter { EmployeeNumber -like "*" -and Enabled -eq $True } -SearchBase "ou=Students,dc=domain,dc=local" | Set-ADUser -Surname "ANYTHING"```

Q. I need to update to the latest version

A. Open Powershell and "cd \scripts\automated_students" then "git pull"