<#

Clever Interim Script
What needs to happen with the Clever files before final upload to Clever?
Some ideas:
- Add the username and password fields
- Add the Title to the Teachers File
- Add the Role to the Staff File.
- Modify to fit your needs. You can strip out the contacts on the students.csv and just use the contacts.csv

#To manage Staff manually you can put them in a Google Sheet named "Clever Staff" in your Automated Students folder. This will allow you to override any automatically generated information.
Get-GDriveSheetId -path "Clever Staff" -InitialIncomingData 'School_id,Staff_id,Staff_email,First_name,Last_name,Department,Title,Username,Password,Role'

#>

<#



#First pull staff that isn't assigned to every building.
$staff = Invoke-SqlQuery -Query "SELECT * FROM staff WHERE `Staff_id` IN (SELECT `staff`.`Staff_id` FROM staff LEFT JOIN teachers_extras ON staff.Staff_id = teachers_extras.Teacher_id GROUP BY `Staff_id` HAVING count(`School_id`) < (select Count(*) from `schools`))" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Is_Advisor,Is_Counselor,Is_Teacher,Timestamp,RowError,RowState,Table,ItemArray,HasErrors
#Then pull staff assigned to all buildings then give School_id of 'district'
$staff += Invoke-SqlQuery -Query "SELECT DISTINCT 'district' AS 'School_id',`staff`.`Staff_id`,`staff`.`Staff_email`,`staff`.`First_name`,`staff`.`Last_name`,`staff`.`Department`,`staff`.`Title`,`staff`.`Username`,`staff`.`Password`,'School Tech Lead' AS `Role` FROM staff WHERE `Staff_id` IN (SELECT `staff`.`Staff_id` FROM staff LEFT JOIN teachers_extras ON staff.Staff_id = teachers_extras.Teacher_id GROUP BY `Staff_id` HAVING count(`School_id`) >= (select Count(*) from `schools`))" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Is_Advisor,Is_Counselor,Is_Teacher,Timestamp,RowError,RowState,Table,ItemArray,HasErrors

#pull additional staff from Google Drive.
if ($GoogleAccount) {
    Write-Host "Info: Pulling Additional Staff from ""Clever Staff"" in Google Drive."
    Get-GDriveSheetId -path "Clever Staff" -InitialIncomingData 'School_id,Staff_id,Staff_email,First_name,Last_name,Department,Title,Username,Password,Role'
    rclone --config C:\Scripts\rclone\rclone.conf copy "google-drive:automated_students/Clever Staff.csv" $env:temp --drive-export-formats CSV
    #only use the file if we are successful copying a new version from Google. Otherwise use the previous file if it exists.
    if ($LASTEXITCODE -eq 0) {
        Copy-Item "$env:temp\Clever Staff.csv" "$currentPath\temp\Clever Staff.csv" -Force
    }
}

if (Test-Path "$currentPath\temp\Clever Staff.csv") {
    $staff += Import-CSV "$currentPath\temp\Clever Staff.csv"
}

$staff | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\staff.csv -Force

$positions = Import-CSV $currentPath\files\teachers_extras.csv | Group-Object -Property Teacher_id -AsHashTable
$staff = Import-CSV "C:\Scripts\automated_students\Clever\staff.csv"
$staff | ForEach-Object {
        $staffid = $PSItem.Staff_id
        if ($positions.$staffid) {
                $title = (($positions.$($staffid))[0]).Complex
                $PSItem.Title = $title

                if ($cleverSTLPositions -contains $title) {
                        $PSItem.Role = "School Tech Lead"
                }
        }
}
$staff | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\staff.csv -Force

#>

<#

$teachers = @()
$teachers += Invoke-SqlQuery -Query 'SELECT 
staff.School_id,
staff.Staff_id AS Teacher_id,
staff.Staff_id AS Teacher_number,
teachers_extras.State_id AS State_teacher_id,
staff.Staff_email AS Teacher_email,
staff.First_name,
'''' AS Middle_name,
staff.Last_name,
Complex AS Title,
staff.Username,
staff.Password
FROM
staff
LEFT JOIN teachers_extras ON staff.Staff_id = teachers_extras.Teacher_id
WHERE
Complex IN (''Principal'',''Counselor'',''Media Specialist'',''Assistant Principal'')' | Select-Object -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors
$teachers += Import-CSV $currentPath\clever\teachers.csv
$teachers | ConvertTo-CSV -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\teachers.csv -Force

#>