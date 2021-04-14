#Requires -Version 7.0

<#

Clever Upload Script
Craig Millsap, Gentry Public Schools, cmillsap@gentrypioneers.com, 3/2021

You must generate a ssh key for Clever and save at .\keys\clever.key then upload the public key at https://schools.clever.com/sync/settings 

If you don't use a private key then you need to specify SkipUploadingFiles and use the x_InterimProcessingClever.ps1
script to upload your files.

Titles are pulled from the Complex or Street_address field in the Staff Catalog.
The role of "School Tech Lead" is applied to any Titles in the $cleverSTLPositions array in settings.ps1
You can add additional Staff and assign roles in the Google Sheet named "Clever Staff"
#Get-GDriveSheetId -path "Clever Staff" -InitialIncomingData 'School_id,Staff_id,Staff_email,First_name,Last_name,Department,Title,Username,Password,Role'

More information about roles for Admins and Staff here:
https://support.clever.com/hc/s/articles/115001733726?language=en_US

#>

Param(
    [Parameter(Mandatory=$false)][switch]$SkipUploadingFiles #Skip Uploading
)

$cleverserver = 'sftp.clever.com'

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

Import-Module SimplySQL
. $currentPath\settings.ps1
. $currentPath\z_functions.ps1
Connect-Database -database $database

if (-Not(Test-Path $currentPath\clever)) { New-Item -ItemType Directory clever }
if (-Not(Test-Path $currentPath\temp)) { New-Item -ItemType Directory temp }
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

#Pull data from PSSQLite and create CSVs for Clever Upload.
Write-Host "Info: Querying data from SQL and building CSVs for Clever..." -ForegroundColor YELLOW
try {
    $students = Invoke-SqlQuery -Query 'SELECT * FROM `students`' -ErrorAction 'STOP' | Select-Object -ExcludeProperty Timestamp,RowError,RowState,Table,ItemArray,HasErrors
    
    $studentEmails = Get-ADUser -Filter { Enabled -eq $True -and EmployeeNumber -like "*" } -Properties EmployeeNumber,EmailAddress,Mail | Group-Object -Property EmployeeNumber -AsHashTable
    
    $students | ForEach-Object {
        
        #Check for mismatched email address.
        $studentId = [string]$PSitem.'Student_id'
        $email = $PSItem.'Student_email'
        
        try {
            $adEmailAddress = ($studentEmails.$studentId).Mail

            if ($email -ne $adEmailAddress) {
                Write-Host "Notify:",$PSItem.Student_id,"is either missing or has an invalid email address on their eSchool record. Pulling from AD." -ForegroundColor Yellow
                $PSItem.Student_email = ($studentEmails.[string]($PSItem.Student_id))[0].EmailAddress #[0] in case of duplicate student ID numbers in AD.
            }
        } catch {}

    }
    
    $students | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\students.csv -Force

    $sections = Invoke-SqlQuery -Query "SELECT * FROM sections WHERE Term_name = $currentTerm" -ErrorAction 'STOP'
    $sections += Invoke-SqlQuery -Query "SELECT * FROM sections WHERE Term_name != $currentTerm" -ErrorAction 'STOP'
    $sections | Select-Object -Property * -ExcludeProperty Timestamp,RowError,RowState,Table,ItemArray,HasErrors | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\sections.csv -Force

    $enrollments = Invoke-SqlQuery -Query "SELECT School_id,Section_id,Student_id FROM schedules" -ErrorAction 'STOP'
    $enrollments | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\enrollments.csv -Force

    if ($clevereSchoolTitleField) {
        $eSchoolTitleQuery = "``teachers_extras``.``$($clevereSchoolTitleField)``"
    } else {
        $eSchoolTitleQuery = ''''''
    }

    Start-SqlTransaction
    if (@('mysql','mariadb') -contains $database.dbtype) {
      Invoke-SqlUpdate -Query "SET SESSION sql_mode=(SELECT REPLACE(@@sql_mode,'ONLY_FULL_GROUP_BY',''));" | Out-Null
    }

    $teachers = Invoke-SqlQuery -Query ('SELECT
    `teachers`.`School_id`,
    `teachers`.`Teacher_id`,
    `teachers`.`Teacher_number`,
    `teachers`.`State_teacher_id`,
    `teachers`.`Teacher_email`,
    `teachers`.`First_name`,
    `teachers`.`Middle_name`,
    `teachers`.`Last_name`,' + $eSchoolTitleQuery + ' AS `Title`,
    '''' AS `Username`,
    '''' AS `Password`
    FROM `teachers`
    LEFT JOIN `teachers_extras` ON `teachers`.`Teacher_id` = `teachers_extras`.`Teacher_id`
    GROUP BY `teachers`.`Teacher_id`')
    Complete-SqlTransaction
    #SELECT * FROM teachers GROUP BY Teacher_id" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp,RowError,RowState,Table,ItemArray,HasErrors
    $teachers | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\teachers.csv -Force

    $schools = Invoke-SqlQuery -Query "SELECT * FROM schools" -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp,RowError,RowState,Table,ItemArray,HasErrors
    $schools | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\schools.csv -Force

    #Contact_id,Student_id,Contact_firstname,Contact_lastname,Contact_type,Contact_phonenumber,Contact_relationship,Contact_phonetype,Contact_email
    #contact_id,student_id,first_name,last_name,type,phone,relationship,phone_type,email
    $contacts = Invoke-SqlQuery -Query 'SELECT
        `contacts`.`Contact_id` AS `Contact_id`,
        `contacts`.`Student_id` AS `Student_id`,
        `contacts`.`Contact_firstname` AS `first_name`,
        `contacts`.`Contact_lastname` AS `last_name`,
        `contacts`.`Contact_type` AS `type`,
        `contacts`.`Contact_phonenumber` AS `phone`,
        `contacts`.`Contact_relationship` AS `relationship`,
        `contacts`.`Contact_phonetype` AS `phone_type`,
        `contacts`.`Contact_email` AS `email`
        FROM `contacts`
        WHERE `Contact_phonetype` IN (''C'',''C1'',''C2'',''CO'')'
    $contacts | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\customcontacts\contacts.csv -Force

    $staff = @()
    if ($cleverSTLPositions) {
        $cleverSTLString = '''' + ($cleverSTLPositions -join ''',''') + ''''
        $staff += Invoke-SqlQuery -Query ('SELECT
        School_id,Staff_id,Staff_email,First_name,Last_name,
        '''' AS Department,
        CASE
            WHEN `Complex` IN (' + $cleverSTLString + ') THEN `Complex`
            WHEN `Street_name` IN (' + $cleverSTLString + ') THEN `Street_name`
        END as Title,
        '''' AS `Username`,
        '''' AS `Password`,
        ''School Tech Lead'' AS `Role`
        FROM staff
        LEFT JOIN `teachers_extras` ON `staff`.`Staff_id` = `teachers_extras`.`Teacher_id`
        WHERE
        Staff_id IN
        (SELECT `Teacher_id` FROM teachers_extras WHERE `Complex` IN (' + $cleverSTLString + ') OR `Street_name` IN (' + $cleverSTLString + '))')
    }

    if ($GoogleAccount) {
        Write-Host "Info: Pulling Additional Staff from ""Clever Staff"" in Google Drive."
        #Do not allow this file to be created again just in case it didn't find it once. It would break stuff.
        #Get-GDriveSheetId -path "Clever Staff" -InitialIncomingData 'School_id,Staff_id,Staff_email,First_name,Last_name,Department,Title,Username,Password,Role' -DoNotCreateFile
        rclone --config C:\Scripts\rclone\rclone.conf copy "google-drive:automated_students/Clever Staff.csv" $env:temp --drive-export-formats CSV
        #only use the file if we are successful copying a new version from Google. Otherwise use the previous file if it exists.
        if ($LASTEXITCODE -eq 0) {
            Copy-Item "$env:temp\Clever Staff.csv" "$currentPath\temp\Clever Staff.csv" -Force
        }
    }
    
    if (Test-Path "$currentPath\temp\Clever Staff.csv") {
        $staff += Import-CSV "$currentPath\temp\Clever Staff.csv"
    }

    #We should import all other staff BECAUSE they should at least have Portal Access to Clever. This should exclude existing things defined in Teachers and Staff and give no roles.
    $existingids = $staff | Select-Object -ExpandProperty Staff_id
    $existingids += $teachers | Select-Object -ExpandProperty Teacher_id
    $existingids = $existingids | Select-Object -Unique

    $staff += Invoke-SqlQuery -Query ('SELECT
    `School_id`,
    `Staff_id`,
    `Staff_email`,
    `First_name`,
    `Last_name`,
    `Department`,'+ $eSchoolTitleQuery + ' AS `Title`,
    `Username`,
    `Password`,
    `Role`
    FROM `staff`
    LEFT JOIN `teachers_extras` ON `staff`.`Staff_id` = `teachers_extras`.`Teacher_id`
    WHERE `Staff_id` NOT IN (' + '''' + ($existingids -join ''',''') + '''' + ')
    ')

    if ($staff.Count -ge 1) {
        #Select-Object -Property * -ExcludeProperty Is_Advisor,Is_Counselor,Is_Teacher,Timestamp,RowError,RowState,Table,ItemArray,HasErrors
        $staff | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\clever\staff.csv -Force
    }

} catch {
    write-host "ERROR: Failed to query SQL and create CSVs. $PSItem" -ForegroundColor RED
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