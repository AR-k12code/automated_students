& c:\scripts\gam\gam.exe ou_and_children "/Students" print suspended > c:\scripts\gam\students_suspended.csv
Start-Process -FilePath "c:\scripts\gam\gam.exe" -ArgumentList 'ou_and_children "/Students" print suspended' -RedirectStandardOutput "c:\scripts\gam\students_suspended.csv" -Wait -NoNewWindow -PassThru

$students = import-csv students_suspended.csv | Where-Object { $PSitem.suspended -eq "True" }

$students | ForEach-Object {
    Start-Process -FilePath "c:\scripts\gam\gam.exe" -ArgumentList "update user $($PSItem.primaryEmail) ou /Students/Disabled" -Wait -NoNewWindow -PassThru
}