# This file is provided as a launcher for the automated_students project.
# The goal being that QuickMode should be launched automatically between the hours of 6AM - 4PM.
# There are School Districts where the full Student Configuration Management process can take over an hour and nobody should have to wait that long for a new account information.

$hour = [int](Get-Date -Format HH)

if ($hour -ge 6 -and $hour -lt 16) {
    Start-Process "pwsh.exe" -ArgumentList "-ExecutionPolicy bypass -File c:\scripts\automated_students\automated_database.ps1 -QuickMode" -NoNewWindow -WorkingDirectory "c:\scripts\automated_students" -Wait
} else {
    Start-Process "pwsh.exe" -ArgumentList "-ExecutionPolicy bypass -File c:\scripts\automated_students\automated_database.ps1" -NoNewWindow -WorkingDirectory "c:\scripts\automated_students" -Wait
}