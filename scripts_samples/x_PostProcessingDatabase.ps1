#Requires -Version 7.0
# This file is called after a successful database build.

Param(
    [Parameter(Mandatory=$false)][switch]$QuickMode #This will only import the Students and Students_extras reports.
)

#archive AD structure to CSV. This is a CYA. We want to make sure we have a copy of the AD structure before we start making changes.
if (-Not(Test-Path $currentPath\archives)) { New-Item -ItemType Directory -Path $currentPath\archives }
Get-ADUser -Filter * -Properties DisplayName,givenName,Surname,EmployeeNumber,EmployeeID,SamAccountName,UserPrincipalName,EmailAddress,DistinguishedName,Enabled,Fax,HomePhone,Office,ObjectGUID,objectSid | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\archives\users_ad_structure_$(get-date -Format MM-dd-yyyy-HH-mm-ss).csv"

#Pull Overrides and Exclusions from Google Drive. Do not create the files because we never want to overwrite or create a new sheet on accident.
# From here we want to launch our students_extras.ps1 script.
if ($QuickMode) {
    Start-Process -FilePath "pwsh.exe" -ArgumentList "-f automated_students.ps1 -QuickMode" -NoNewWindow -Wait
} else {
    Start-Process -FilePath "pwsh.exe" -ArgumentList "-f automated_students.ps1" -NoNewWindow -Wait
}