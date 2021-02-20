# Clever Interim Script
# What needs to happen with the Clever files before final upload to Clever?


#Upload Clever Files to ScriptK12
# VARIABLES
$scriptusername = 'gentry'
$scriptserver = 'sample.serverhere.com'
$scriptpath = '/srv/share/gentry/clever'

try {
    Write-Host "Uploading files to ScriptK12..." -ForegroundColor YELLOW
    Start-Process -FilePath "$currentPath\bin\pscp.exe" `
        -ArgumentList "-i $currentpath\keys\scriptk12.ppk $currentpath\files\schools.csv $currentpath\files\teachers.csv $currentpath\files\students.csv $currentpath\files\sections.csv $currentpath\files\enrollments.csv $($scriptusername)@$($scriptserver):$($scriptpath)" -PassThru -Wait -NoNewWindow
} catch {
        write-Host "Failed to properly upload files to Scriptk12." -ForegroundColor RED
}
