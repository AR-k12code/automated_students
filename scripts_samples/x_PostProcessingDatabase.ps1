#Requires -Version 7.0
# This file is called after a successful database build.

Param(
    [Parameter(Mandatory=$false)][switch]$QuickMode #This will only import the Students and Students_extras reports.
)

#archive AD structure to CSV.
if (-Not(Test-Path $currentPath\archives)) { New-Item -ItemType Directory -Path $currentPath\archives }
Get-ADUser -Filter * -Properties DisplayName,givenName,Surname,EmployeeNumber,EmployeeID,SamAccountName,UserPrincipalName,EmailAddress,DistinguishedName,Enabled,Fax,HomePhone,Office,ObjectGUID,objectSid | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\archives\users_ad_structure_$(get-date -Format MM-dd-yyyy-HH-mm-ss).csv"

#Pull Overrides and Exclusions from Google Drive. Do not create the files because we never want to overwrite or create a new sheet on accident.
try {
    $overrides = Get-GDriveSheetId -path "overrides" -DoNotCreateFile
    #Get-GDriveSheetId -path "overrides" -InitialIncomingData "Student_id,First_name,Last_name,Middle_Initial,Grade"
    rclone --config C:\Scripts\rclone\rclone.conf copy google-drive:automated_students/overrides.csv $env:temp --drive-export-formats CSV 
    if ($LASTEXITCODE -eq 0) { Copy-Item $env:temp\overrides.csv c:\scripts\automated_students\overrides.csv -Force }
    $exclusions = Get-GDriveSheetId -path "exclusions" -DoNotCreateFile
    #Get-GDriveSheetId -path "exclusions" -InitialIncomingData "Student_id,First_name,Last_name"
    rclone --config C:\Scripts\rclone\rclone.conf copy google-drive:automated_students/exclusions.csv $env:temp --drive-export-formats CSV 
    if ($LASTEXITCODE -eq 0) { Copy-Item $env:temp\exclusions.csv c:\scripts\automated_students\exclusions.csv -Force }
} catch {}

try {
    Write-Host "Info: Backing up Database to Google Drive."
    if (@('mysql','mariadb') -contains $database.dbtype) {
       & mysqldump.exe -u automated_students -p"$($database.password)" --databases "$($database.dbname)" --result-file=database_backup.sql
       rclone --config c:\scripts\rclone\rclone.conf copy database_backup.sql google-drive:dbbackups\ -v
    } else {
        rclone --config c:\scripts\rclone\rclone.conf copy $database google-drive:dbbackups\ -v
        if ($LASTEXITCODE -ne 0) { Throw 'Failed to backup Database to Google Drive.' }
    }
} catch { $PSItem }

# From here we want to launch our students_extras.ps1 script.
if ($QuickMode) {
    Start-Process -FilePath "pwsh.exe" -ArgumentList "-f automated_students.ps1 -QuickMode" -NoNewWindow -Wait
} else {
    Start-Process -FilePath "pwsh.exe" -ArgumentList "-f automated_students.ps1" -NoNewWindow -Wait
}

#Clever
#Moved to a Syncronous Post Automated Students Task.
#Start-Process -FilePath "pwsh.exe" -ArgumentList "-f clever.ps1" -NoNewWindow -Wait
