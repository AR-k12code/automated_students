#Custom Get-NewPassword to override the hard coded one.
function Get-NewPassword($student) {

    #$student comes in with the following properties from the students table:
    #Contact_email,Contact_name,Contact_phone,Contact_relationship,Contact_type,DOB,Ell_status,First_name,Frl_status,Gender,Grade,Hispanic_Latino,Home_language,Iep_status,Last_name,Middle_name,Password,Race,School_id,State_id,Student_city,Student_email,Student_id,Student_number,Student_state,Student_street,Student_zip

    #How to create a custom password for individual buildings.
    if (@(1,2,3,4) -contains $student.School_id) {
        $newPassword = '1Pioneers' + ([string]$($student.Student_id)).Substring(([string]$($student.Student_id)).length - 4, 4)
    } elseif (@(5,6) -contains $student.School_id) {
        $newPassword = '1Bulldogs' + ([string]$($student.Student_id)).Substring(([string]$($student.Student_id)).length - 4, 4)
    } else {
        $newPassword = [string]$(Get-RandomCharacters 8 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNOPQRSTUVWXYZ123456789!.$#%&*<>') #no l,o, or 0. So there is no confusion.
    }

    return $newPassword
}

