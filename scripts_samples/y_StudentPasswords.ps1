$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

Import-Module PSSQLite
. $currentPath\settings.ps1

$gdrivefolder = "1-6NYTemPPZ7dmY00000000000000000" #google drive folder that holds the spreadsheets.
$blankspreadsheet = "1h74-R4eOQ8XharKxCXHMRqSADHcYLLlT4FJiyPC-eBQ" #any blank spreadsheet that will never get deleted.
$buildings = '(15,13,16)' #valid buildings you want to pull passwords for and share to any teacher they are scheduled to.

if (-Not(Test-Path "$currentPath\passwords\teachers\")) {
    New-Item -Path "$currentPath\passwords\teachers\" -ItemType Directory -Force
}

$q = "SELECT DISTINCT teachers.[Teacher_id],[Teacher_email],students.[Student_id],students.[First_name],students.[Last_name],students.[Grade],students.[Student_email],passwords.[Student_password] FROM teachers `
INNER JOIN sections ON teachers.[Teacher_id] = sections.[Teacher_id] `
INNER JOIN enrollments ON sections.[Section_id] = enrollments.[Section_id] `
INNER JOIN students ON enrollments.[Student_id] = students.[Student_id] `
INNER JOIN passwords ON students.[Student_id] = passwords.[Student_id] `
WHERE teachers.[School_id] IN $buildings AND students.[School_id] IN $buildings `
ORDER BY students.[Last_name]"

$results = Invoke-SqliteQuery -DataSource $database -Query $q | Group-Object -Property Teacher_id

#get files from google drive to find ID's to update.
try {
    $files = rclone --config C:\Scripts\rclone\rclone.conf lsjson "google-drive:Projects/Student Default Passwords/" -R
    $files = $files | ConvertFrom-Json
} catch {
    write-host "Error listing files from Google Drive."
    exit(1)
}

#$header = "Student_id,First_name,Last_name,Grade,Student_email,Password"
$results | ForEach-Object {

    $teacher_email = $PSItem.Group | Select-Object -First 1 -ExpandProperty 'Teacher_email'
    $teacher_id = $PSItem.Group | Select-Object -First 1 -ExpandProperty 'Teacher_id'
    $teacherfileid = "$($($teacher_email.Split('@'))[0])-$($teacher_id)"

    if ($teacher_email -eq '') { continue }
    
    $students = $PSItem.Group | Sort-Object -Property "Student_id" -Unique | Sort-Object -Property 'Last_name' | Select-Object -Property 'Student_id','First_name','Last_name','Grade','Student_email','Student_password'
    
    $filename = "Student Passwords for $teacherfileid"

    $lines = "Student_id,First_name,Last_name,Grade,Student_email,Password`r`n"
    
    $students | ForEach-Object {
        $lines += "$($PSItem.'Student_id'),$($PSItem.'First_name'),$($PSItem.'Last_name'),$($PSItem.'Grade'),$($PSItem.'Student_email'),$($PSItem.'Student_password')`r`n"
    }

    Out-File -FilePath ".\passwords\teachers\$($filename).csv" -InputObject $lines -NoNewline -Force

    if ($files | Where-Object { $PSitem.'Name' -like "*$teacherfileid*" }) {
        $gdrivefileid = $files | Where-Object { $PSitem.'Name' -like "*$teacherfileid*" } | Select-Object -ExpandProperty ID -First 1
        gam user technology update drivefile id $gdrivefileid localfile "$currentPath\passwords\teachers\$($filename).csv"
    } else {
        write-host "Info: No existing file match found in google drive. Creating new file and uploading contents."
        try {
            #the only way to get the csv's to upload as a google sheet is to already have a sheet in there. This makes a copy of a blank sheet.
            $newfileinfo = gam user technology copy drivefile $blankspreadsheet newfilename "$filename" parentid $gdrivefolder
            $newfileid = $($($($($newfileinfo | Select-String -Pattern "Drive File ID:") -split ',' | Select-String -Pattern "Drive File ID:") -split ":")[1]).Trim()
            gam user technology update drivefile id $newfileid localfile "$currentPath\passwords\teachers\$($filename).csv" newfilename "$filename"
            gam user technology add drivefileacl $newfileid user $teacher_email role reader sendemail
        } catch {
            write-host "Unable to create new sheets."
            exit(1)
        }
        
    }

}
