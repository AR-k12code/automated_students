#Requires -Version 7.0
$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

Import-Module PSSQLite
. $currentPath\settings.ps1

#archive AD User structure to CSV before any work is done.
if (-Not(Test-Path $currentPath\archives)) { New-Item -ItemType Directory -Path $currentPath\archives }
Get-ADUser -Filter * -Properties DisplayName,givenName,Surname,EmployeeNumber,EmployeeID,SamAccountName,UserPrincipalName,mail,DistinguishedName,Enabled,Fax,HomePhone,Office,ObjectGUID,objectSid | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\archives\users_ad_structure_$(get-date -Format MM-dd-yyyy-HH-mm-ss).csv"

# This file is called after a successful PSSQLite database build.
# From here we want to launch our students_extras.ps1 script.
Start-Process -FilePath "pwsh.exe" -ArgumentList "-f automated_students.ps1" -NoNewWindow -Wait

#Clever
Start-Process -FilePath "pwsh.exe" -ArgumentList "-f clever.ps1" -NoNewWindow -Wait

###############################################################
# Everything below this line is customized for your district. #
###############################################################

######################################################################################
# Create Spreadsheets for our PVLA and Onsite Students and put them in Google Drive. #
######################################################################################

if (-Not(Test-Path $currentPath\temp)) { New-Item -ItemType Directory -Path $currentPath\temp }

$q = "SELECT * FROM students 
    WHERE Student_id 
    NOT IN (select student_id from students_extras WHERE students_extras.Student_houseteam = 'PVLA ')
    ORDER BY School_id,Last_name"

$onsitestudents = Invoke-SqliteQuery -DataSource $database -Query $q

$onsitestudents | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File .\temp\onsite.csv -Force
gam user technology@gentrypioneers.com update drivefile id 1So-zu3ptU2sn0ItMUafPziuorBd5pcyeUvou50opXt0 localfile .\temp\onsite.csv newfilename "Onsite Students"

$q = "SELECT * FROM students
    WHERE Student_id
    IN (SELECT student_id FROM students_extras WHERE students_extras.Student_houseteam = 'PVLA ')
    ORDER BY School_id,Last_name"

$pvlastudents = Invoke-SqliteQuery -DataSource $database -Query $q

$pvlastudents | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File .\temp\pvla.csv -Force
gam user technology@gentrypioneers.com update drivefile id 1bc7MaVFKBuyh7GfcWQLq2HhvzxdPnqUbPhPaOtnkV34 localfile .\temp\pvla.csv newfilename "PVLA Students"

#############################################################################
# Create PVLA Student Group in AD. Set WbemPath to the moderators manually. #
#############################################################################
$q = 'SELECT students.School_id,students.Student_email,students.Grade FROM students_extras 
INNER JOIN students ON students_extras.Student_id = students.student_id 
WHERE Student_houseteam = "PVLA "'

Invoke-SqliteQuery -DataSource $database -Query $q | Select-Object -ExpandProperty Student_email | ForEach-Object {
    Add-ADGroupMember -Identity students-pvla -Members (Get-ADUser -Filter "UserPrincipalName -eq ""$PSItem""")
}

###########################################################
# Upload School Specific CSVs for Gradebook Exports in GC #
###########################################################
$gdriveuser = "technology@gentrypioneers.com"
$parentidclever = "1kqF4OPxz2ahf54H00000000000000000"
Invoke-SqliteQuery -DataSource $database -Query "SELECT Student_email,Student_id FROM students WHERE School_id = 16" | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\files\GPS Students GC Match.csv" -Force
gam user $gdriveuser update drivefile id 1LF5pZlhkGWcJ7nkyP9twj6MVri00000000000000000 localfile "$currentPath\files\GPS Students GC Match.csv" newfilename "GPS Students GC Match" parentid $parentidclever
Invoke-SqliteQuery -DataSource $database -Query "SELECT Student_email,Student_id FROM students WHERE School_id = 13" | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\files\GIS Students GC Match.csv" -Force
gam user $gdriveuser update drivefile id 11FQXXfmLbPX0JJ_BNtRNh45kCV00000000000000000 localfile "$currentPath\files\GIS Students GC Match.csv" newfilename "GIS Students GC Match" parentid $parentidclever
Invoke-SqliteQuery -DataSource $database -Query "SELECT Student_email,Student_id FROM students WHERE School_id = 15" | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\files\GMS Students GC Match.csv" -Force
gam user $gdriveuser update drivefile id 16sxRGVMEGx-XZFxcDU50kUZ6Ho00000000000000000 localfile "$currentPath\files\GMS Students GC Match.csv" newfilename "GMS Students GC Match" parentid $parentidclever
Invoke-SqliteQuery -DataSource $database -Query "SELECT Student_email,Student_id FROM students WHERE School_id = 703" | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\files\GHSCC Students GC Match.csv" -Force
gam user $gdriveuser update drivefile id 1fNXaOK6NR3rXwnP-79G1fiM90j00000000000000000 localfile "$currentPath\files\GHSCC Students GC Match.csv" newfilename "GHSCC Students GC Match" parentid $parentidclever
