#This script is called for existing users. You have access to all these variables for the student from the CSV.
#$studentId,$firstName,$lastName,$fullName,$grade,$gender,$buildingNumber,$buildingShortName,$gradyr,$emailDomain
#$password,$username,$principalName,$emailAddress,$homeDirRoot,$homeDirGradYR,$homeDir,$ou

#Be aware of $staging. This script is called either way.

#$existingAccount - contains the current student AD account information. All properties from Active Directory. Used for comparisons for name changes etc.
#$existingAccountGUID - references the GUID of the existing user account regardless of name chnages that might have happenend prior to this script launching.

#Set SIS Number to EmployeeID field in AD.
# if (-Not($Staging)) {
#     if (($existingAccount.EmployeeID -ne $PSItem.State_id) -and ($PSItem.State_id -gt 1)) {
#         Set-ADUser -Identity $existingAccountGUID -EmployeeID $PSItem.State_id
#     }
# }

#Custom Home Directory Drive and Logon Script
# if (($existingAccount.HomeDrive -ne "S:") -or ($existingAccount.ScriptPath -ne "logon.bat")) {
#     Set-ADUser -Identity $existingAccountGUID -HomeDrive "S:" -ScriptPath "logon.bat"
# }

#Students Moved into the buildings 16 and 13 shouldn't be able to change thier password.
#The additional if statements are to limit the number of changes. Otherwise we overwhelm our DCs.
# if (@(16,13) -contains $buildingNumber) {
#     if (($existingAccount.PasswordNeverExpires -eq $True) -or ($existingAccount.CannotChangePassword -eq $False)) {
#         Set-AdUser -Identity $existingAccountGUID -ChangePasswordAtLogon $False
#         Set-ADAccountControl -Identity $existingAccountGUID -CannotChangePassword $True
#     }
# } else {
#     if (($existingAccount.PasswordNeverExpires -eq $False) -or ($existingAccount.CannotChangePassword -eq $True)) {
#         Set-AdUser -Identity $existingAccountGUID -ChangePasswordAtLogon $False
#         Set-ADAccountControl -Identity $existingAccountGUID -CannotChangePassword $False
#     }
# }

#Lets reset all student accounts passwords and save to a building level CSV.
if ($ResetAllPasswords) {

    write-host "Updating password for $fullName, $gradyr, $buildingNumber"

    #Insert new password into the database.
    Invoke-SqlQuery -Query "REPLACE INTO passwords (Student_id, Student_password, HAC_passwordset, Timestamp) VALUES ($studentid,""$password"",NULL,$timestamp)"

    if ($staging) {     
        Set-AdAccountPassword -Identity $existingAccountGUID -Reset -NewPassword (ConvertTo-SecureString "$password" -AsPlainText -Force) -WhatIf
        Set-AdUser -Identity $existingAccountGUID -ChangePasswordAtLogon $False -WhatIf
        Set-ADAccountControl -Identity $existingAccountGUID -CannotChangePassword $False -WhatIf
    } else {
        Set-AdAccountPassword -Identity $existingAccountGUID -Reset -NewPassword (ConvertTo-SecureString "$password" -AsPlainText -Force)
        Set-AdUser -Identity $existingAccountGUID -ChangePasswordAtLogon $False
        Set-ADAccountControl -Identity $existingAccountGUID -CannotChangePassword $False
    }
    
}

# Sample of adding an expiration in the future. If you turn off disabling accounts this is a good way to allow for accounts to be automatically disabled in the future. You need a
# post processing script to disable expired accounts.
# Set-ADAccountExpiration -Identity $existingAccountGUID -DateTime (Get-Date).AddDays(3)