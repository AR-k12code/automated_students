#This is stuff we want to do after running our Student Automation Script.

#Move Disabled Students to the Disabled OU. This is written for AD structures 1 & 2. It will not work for #3. Someone else can write that and merge it back into the main script.
Get-ADUser -Filter { Enabled -eq $False } -SearchBase "ou=Students,$($domain)" -Properties EmployeeNumber |
    Where-Object { $PSItem.DistinguishedName -notlike "*,ou=Disabled,ou=Students,$($domain)" } | ForEach-Object {
        if ($excludedStudentIDs -contains $PSitem.EmployeeNumber) { return }
        Write-Host "Info: Moving disabled account $($PSItem.DistinguishedName) to Disabled OU."
        Move-ADObject -Identity $PSitem.ObjectGUID -TargetPath "ou=Disabled,ou=Students,$($domain)"
    }

#create a students.csv for each school that is updated every time this script runs.
if (-Not(Test-Path $currentPath\exports\students)) { New-Item -ItemType Directory -Path "$currentPath\exports\students" -Force }
$validschools.Keys | ForEach-Object {
    Invoke-SqliteQuery -DataSource $database -Query "select * from students WHERE School_id = $PSItem" | Select-Object -Property * -ExcludeProperty Timestamp | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\exports\students\students-$($validschools.$PSItem).csv" -Force
}

#drivefile ids are obfuscated to security. You will need to create your own blank spreadsheets, then copy them to the correct lines.

#For me I upload the data to Google Drive using GAM and rclone. That was I can inspect the files directly from my workstation.
write-host "Info: Updating CSV files in Google Drive..."

#For us its important to have updated files in Google Drive.
$gdriveuser = "technology@gentrypioneers.com"

#You should manually create these sheets in Google Drive under the user below. Get the file id and update that specific file using gam update.
#clever files converted to google sheets
$parentidclever = "1kqF4OPxz2ahf5000000000000000000000000"
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 10x7www5OzhoSj6ePUZo000000000000000000000000 localfile "$currentPath\clever\students.csv" newfilename "students" parentid $parentidclever #tdsheet id:555791402
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 17Kfs-losMCTt3xXjW7W000000000000000000000000 localfile "$currentPath\clever\teachers.csv" newfilename "teachers" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1q5MYDlhJfken-bZFYsN000000000000000000000000 localfile "$currentPath\clever\schools.csv" newfilename "schools" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1BvIJZhagUXktNLx0Gme000000000000000000000000 localfile "$currentPath\clever\sections.csv" newfilename "sections" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1sKQv-snsbM_DFiZ_pu2000000000000000000000000 localfile "$currentPath\clever\enrollments.csv" newfilename "enrollments" parentid $parentidclever

#clever files kept as CSV
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 10uUosFiPrSYyfPEo5a0000000000000000000000000 localfile "$currentPath\clever\students.csv" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1qsHU5vvUKrBfV1LZqvN000000000000000000000000 localfile "$currentPath\clever\teachers.csv" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1pdTib1jxHKyCYhtsbII000000000000000000000000 localfile "$currentPath\clever\schools.csv" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1lsvVlw_8u8JW9f-EdZG000000000000000000000000 localfile "$currentPath\clever\sections.csv" parentid $parentidclever
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1wi2TrPTAO_i3K9-jsgD000000000000000000000000 localfile "$currentPath\clever\enrollments.csv" parentid $parentidclever

#student password files
$parentidpasswords = "1Q-o4PH1MTvceEAnJ6huSrJVjfcnHzTRn"
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 11zjEB9EB9OAzXkUIGd0000000000000000000000000 localfile "$currentPath\passwords\GPS-passwords.csv" newfilename "GPS Student Passwords" parentid $parentidpasswords
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1kX25fjwyNHwM5eX6QbV000000000000000000000000 localfile "$currentPath\passwords\GIS-passwords.csv" newfilename "GIS Student Passwords" parentid $parentidpasswords
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1cvOuMZb4HelAuIxp3yc000000000000000000000000 localfile "$currentPath\passwords\GMS-passwords.csv" newfilename "GMS Student Passwords" parentid $parentidpasswords
& $currentPath\..\gam\gam.exe user $gdriveuser update drivefile id 1HlfHJ6Tf80-34W2ppTz000000000000000000000000 localfile "$currentPath\passwords\GHSCC-passwords.csv" newfilename "GHSCC Student Passwords" parentid $parentidpasswords

#Clear old GCDS locks. This is a constant problem. This should be the only place this runs now.
Remove-Item -Path "$env:userprofile\syncState\*.lock" -Force
#This sync will rename accounts based on an existing key link.
Start-Process -FilePath 'c:\Program Files\Google Cloud Directory Sync\sync-cmd.exe' -ArgumentList '-a -o -c c:\scripts\gads\gentry-users' -wait -NoNewWindow; Start-Sleep -Seconds 15
#This forced sync will take over accounts that gam created and move them to the correct OU.
Start-Process -FilePath 'c:\Program Files\Google Cloud Directory Sync\sync-cmd.exe' -ArgumentList '-a -o -c c:\scripts\gads\gentry-users -f' -wait -NoNewWindow; Start-Sleep -Seconds 15
#Create and sync groups.
Start-Process -FilePath 'c:\Program Files\Google Cloud Directory Sync\sync-cmd.exe' -ArgumentList '-a -o -c c:\scripts\gads\gentry-groups' -wait -NoNewWindow; Start-Sleep -Seconds 15

#######################################
#Moderation Rules of Groups
#######################################
$validschools.Values | ForEach-Object {

    Get-ADGroup -SearchBase "ou=Students,dc=Gentry,dc=local" -Filter " name -like ""students-$($PSItem)*"" -and mail -like ""*@gentrystudents.com""" -Properties mail | ForEach-Object {
    
        $groupemail = $PSItem.'mail'
        $groupinfo = gam info group $groupemail
        Write-Output "Checking moderation settings for $groupemail ..."

        if (!($groupinfo | Select-String -Pattern "membersCanPostAsTheGroup: false")) {
            gam update group $groupemail members_can_post_as_the_group false
        }
        if (!($groupinfo | Select-String -Pattern "spamModerationLevel: MODERATE")) {
            gam update group $groupemail spam_moderation_level moderate
        }

        if (!($groupinfo | Select-String -Pattern "whoCanPostMessage: ALL_MANAGERS_CAN_POST")) {
            gam update group $groupemail who_can_post_message all_managers_can_post
        }
    
        #with whoCanPostMessage set to Managers then we shouldn't need to moderate the messages.
        # if (!($groupinfo | Select-String -Pattern "messageModerationLevel: MODERATE_ALL_MESSAGES")) {
        #     gam update group $groupemail message_moderation_level moderate_all_messages
        # }
    }
}

#If there are emails in eSchool to fix or populate then we need to run our eSchoolUpload. Be careful. Exclusions without an email in eSchool can cause a loop.
if (($studentsCSV | Where-Object { $PSItem.'Student_email' -notlike "*@gentrystudents.com" } | Measure-Object).count -ge 1) {
    write-host "Info: Detected students in eSchool who do not have an email address assigned. Running eSchoolUpload and generating HAC Logins."
    Start-Process -FilePath "pwsh.exe" -ArgumentList "-f c:\scripts\eSchoolUpload\Gentry-UpdateEmailAddresses.ps1" -NoNewWindow -Wait -WorkingDirectory "C:\scripts\eSchoolUpload"
    write-host "Info: Waiting 3 minutes for eSchool to process all changes then running the automated_students script again."
    Start-Sleep -Seconds 180  #Wait 3 minutes to let eSchool process all changes

    #if this is true we need to run the sqlite script again to verify the accounts were updated.
    Write-Host "Info: Running this script again to verify that eSchool was updated."
    Start-Process -FilePath "pwsh.exe" -ArgumentList "-f c:\scripts\automated_students\automated_sqlite.ps1"
    exit
}

###########################################################
# HAC Passwords on VPS
###########################################################
write-host "Info: Updating HAC Passwords on VPS."
$q = 'SELECT 
    students.Student_email,students_extras.Student_hacpassword 
    from students_extras
    LEFT JOIN students on students_extras.Student_id = students.Student_id 
    WHERE students_extras.Student_hacpassword != '''''
$hacpasswords = Invoke-SqliteQuery -Database $database -Query $q | ConvertTo-Json

$user = ''
$pass = ''
$uri = 'https://hac.yourdomainname.com/update.php';
$pair = "$($user):$($pass)"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
$basicAuthValue = "Basic $encodedCreds"
$Headers = @{
    Authorization = $basicAuthValue
}

Invoke-RestMethod -Uri $uri -Body $hacpasswords -Method Put -Headers $Headers -ContentType 'application/json'

#Running Student Import File Scripts
write-host "Info: Running student-imports script."
Start-Process -FilePath "pwsh.exe" -ArgumentList "-f c:\scripts\student-imports\student-import-files.ps1" -NoNewWindow -Wait -WorkingDirectory "c:\scripts\student-imports\"

write-host "Info: Sync Managed Google Classrooms"
Start-Process -FilePath "pwsh.exe" -ArgumentList "-f c:\scripts\classroom\classrooms.ps1" -NoNewWindow -Wait -WorkingDirectory "C:\scripts\classroom"

write-host "Info: Update default password spreadsheets"
Start-Process -FilePath "pwsh.exe" -ArgumentList "-f y_StudentPasswords.ps1" -NoNewWindow -Wait
