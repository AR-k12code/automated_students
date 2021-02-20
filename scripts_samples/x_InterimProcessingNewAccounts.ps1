#This script is called for new users. You have access to all these variables for the student from the CSV.
#$studentId,$firstName,$lastName,$fullName,$grade,$gender,$buildingNumber,$buildingShortName,$gradyr,$emailDomain
#$password,$username,$principalName,$emailAddress,$homeDirRoot,$homeDirGradYR,$homeDir,$ou

#In my district students in buildings 16 and 13 are not allowed to change their passwords.
#It must be done by somebody in the Student Management Group for that Building
# if (@(16,13) -contains $buildingNumber) {
#     Set-AdUser -Identity $username -ChangePasswordAtLogon $False
#     Set-ADAccountControl -Identity $username -CannotChangePassword $True
# }
