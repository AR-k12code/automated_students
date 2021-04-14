$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
if (-Not($domain)) { $domain = (Get-ADDomain).DistinguishedName }
if (-Not(Get-Module -Name SimplySQL)) { Try { Import-Module -Name SimplySQL } catch {} }

#needed eSchool Information for CognosDownloder script.
$eSchoolUsername = '0401cmillsap'
$eSchooldsn = 'gentrysms'
$database = "$($currentPath)\$($eSchooldsn).sqlite3" #don't touch this

# $database = @{
#   dbtype = 'mysql'
#   hostname = 'localhost'
#   dbname = 'gentrysms'
#   username = 'automated_students'
#   password = '*zRpiAfbUZE[z!PG'
# }

$UseSharedCognosFolder = $True #otherwise you copy the automation folder to your own and disable this.

#Clever
$cleverusername = 'clever-random-usenrame'

#These are positions that should have the School Tech Lead Role in Clever. These are automatically pulled from eSchool Staff Catalog Street Address or Complex fields. Can be manually managed in Google Sheets as well.
$cleverSTLPositions = @('Superintendent','Assistant Superintendent','Principal','Assistant Principal','Counselor','Media Specialist')
$clevereSchoolTitleField = 'Complex' #Street_address,Complex, or $False

#This account should be configured with rclone and gam
$GoogleAccount = "gps_automation@mydomain.com"

#Keep log files for x number of days.
$keeplogfiledays = 30

#Minimum student count to expect. This is just a checks and balances to make sure we count valid students and didn't get a bad CSV coming in.
$minStuCount = 1450

#Maximum number of new or disabled student accounts expected in a single day. For summer rollover this needs to be changed. Modifying existing accounts does not apply here.
$maxChanges = 10

#Define valid schools with their School_id numbers and shortname from schools.csv. Shortname must be less than 9 characters
$validschools = @{
    703 = 'GHSCC'
    15 = 'GMS'
    16 = 'GPS'
    13 = 'GIS'
}


#These grades will not be imported into the database when built. Example @('Prekindergarten','PK','Kindergarten','KG')
$ExcludeGrades = @('Prekindergarten','PK')

#Student ID Prefix (we use district LEA number). If you don't have a standarized prefix then leave blank.
$studentIDPrefix = '4010'

#Student email address domain. This can be set per building number.
$stuEmailDomain = @{
    #15 = '@facultydomain.com'
    #1234 = '@facultydomain.com' #you can specify email address domains separately for each school by matching School_id
    default = '@studentdomain.com'
}

$teacherDomain = "@facultydomain.com"

#Define AD Structure
#1 = STUDENTS/GRADYR
#2 = STUDENTS/SCHOOL/GRADYR
#3 = SCHOOL/STUDENTS/GRADYR
#4 = STUDENTS/GRADE
#5 = STUDENTS/SCHOOL/GRADE
#6 = SCHOOL/STUDENTS/GRADE
$adStructure = 2

#Define student username template
#1 [firstname][lastname]
#2 [firstname][lastname][last2ofgradyr]
#3 [firstname].[lastname]
#4 [firstname].[lastname][last2ofgradyr]
#5 [lastname].[firstname]
#6 [lastname].[firstintial].[last2ofgradyr]
#7 [firstname].[first3oflastname][last2ofgradyr]
#8 [first5oflastname][first5offirstname][last2ofgradyr]
$stuTemplate = 1

#Use DinoPass API for a simple password?
$UseDinoPassSimple = $False
$UseDinoPassStrong = $False

#Special OU's that we don't want to move students out of. For Example: Blocking Internet or Allowing Email
$specialOUs = @('Restricted','Blocked','Specials','GT','Allowed','EAST','Excluded')

#Disable Renaming Accounts This overrides the $stuGradYrRename
$DisableRenamingAccounts = $False

#If you use 2,4, or 6 for the student username template do you want to rename students accounts as well if their gradyr changes?
$stuGradYrRename = $True

#use Nickname as the firstname field?
$useNickname = $False

#If a students last name hyphenated do you want to use only the first part? Example: Carlos-Junior would result in only Carlos.
$useFirstHypenatedLastName = $False

#If a students last name has a space do you want to use only the first part? Example: "Carlos Junior" would result in only Carlos.
$useFirstSpacedLastName = $False

#Define servers for home directories based on school id number. Leave blank for no UNC home dir.
#Needs a trailing slash. Home Directories will have the SCHOOLSHORTNAME\GRADYR\USERNAME appended.
#These shares should already exist. They should have FULL ACCESS for EVERYONE or STUDENTS Group on the SHARE.
#NTFS Permissions should have enheritance disabled. With only Domain Admins and SYSTEM with FULL CONTROL to keep
#students from accessing each others folders.
$homeDirectoryRoot = @{
    #703 = '\\gentry1.gentry.local\student-homes$\'
    #15 = '\\gentry3.gentry.local\students$\'
    #13 = '\\gadm1.gentry.local\students\'
}

#What permissions should the Management Group have on student home directories? ReadAndExecute, FullControl, Modify?
$stuManagementHomeDirPerms = 'ReadAndExecute'

#Do you want a random 8 character password for new passwords?
#If not, the default of a 5 letter word, 1 special character, and the last 4 will be used.
#You can override this by creating a z_functionscustom.ps1 with a function called Get-NewPassword($student) that returns what you want.
#$RandomPassword = $True

#send email notications?
$sendMailNotifications = $True
$smtpAuth = $False
$smtpPasswordFile = "C:\Scripts\emailpw.txt"
$sendMailToEmail = 'technology@facultydomain.com'
$sendMailFrom = 'technology@facultydomain.com'
$sendMailHost = 'smtp-relay.gmail.com'
#$sendMailHost = 'smtp.gmail.com'
$sendMailPort = 587

#whom do we notify at each campus?
$notifyEmails = @{
    703 = @('ablanchard@facultydomain.com','kpipkin@facultydomain.com','jfolker@facultydomain.com')
    15 = @('jbrown@facultydomain.com','ddavenport@facultydomain.com','ksmartt@facultydomain.com')
    16 = @('bharrison@facultydomain.com','sselvidge@facultydomain.com','dheinen@facultydomain.com')
    13 = @('aedwards@facultydomain.com','ndover@facultydomain.com','nphilpott@facultydomain.com')
}


#Options are {FULLNAME},{USEREMAILADDRESS},{PASSWORD},{GRADE},{SCHOOLSHORTNAME},{STUDENTID}
$existingAccountMessage = '
{SCHOOLSHORTNAME} has a new student in grade {GRADE}!

An account for {USEREMAILADDRESS} that was previously disabled has been enabled.

Their new assigned password is {PASSWORD}

Please provide this information to the student or their homeroom teacher. 

Thank you,

Technology Department'

#Options are {FULLNAME},{USEREMAILADDRESS},{OLDUSEREMAILADDRESS},{GRADE},{SCHOOLSHORTNAME},{STUDENTID}
$nameChangeMessage = '
{SCHOOLSHORTNAME} has a name change in {GRADE} grade!

A name change has happened for {FULLNAME}. Account has changed from {OLDUSEREMAILADDRESS} to {USEREMAILADDRESS}.

Please provide this information to the student or their homeroom teacher. 

Thank you,

Technology Department'

#Options are {FULLNAME},{USEREMAILADDRESS},{PASSWORD},{GRADE},{SCHOOLSHORTNAME},{STUDENTID}
$newAccountMessage = '
{SCHOOLSHORTNAME} has a new student in grade {GRADE}!

An account account for {USEREMAILADDRESS} has been created.

Their new assigned password is {PASSWORD}

Please provide this information to the student or their homeroom teacher. 

Thank you,

Technology Department'

#Who are the owners of the groups in G-Suite so only they may email the groups and students don't SPAM distribution lists?
#This is by building short name (above), grade, or building-grade. If you have multiple schools with Grade 5 then don't use just grade. Use the Building-Grade. IE. 'GHS-5' = @()
$studentEmailGroupOwners = @{
    'GPS' = @('vgroomer@facultydomain.com','dheinen@facultydomain.com','noreply@facultydomain.com')
    'GIS' = @('kneal@facultydomain.com','ndover@facultydomain.com','aedwards@facultydomain.com','noreply@facultydomain.com')
    'GMS' = @('lcozens@facultydomain.com','ksmartt@facultydomain.com','ddavenport@facultydomain.com','jbrown@facultydomain.com','noreply@facultydomain.com')
    'GHSCC' = @('bharper@facultydomain.com','blittle@facultydomain.com','kpipkin@facultydomain.com','jfolker@facultydomain.com','ehodges@facultydomain.com','ablanchard@facultydomain.com','noreply@facultydomain.com')
    'K' = @()
    '1' = @()
    '2' = @()
    '3' = @()
    '4' = @()
    '5' = @()
    'GIS-5' = @()
    '6' = @()
    '7' = @()
    '8' = @()
    '9' = @('jcampbell@facultydomain.com') #Freshman Class Sponsor
    '10' = @()
    '11' = @('rtingley@facultydomain.com') #Junior Class Sponsor
    '12' = @()
    'SS' = @()
    'students' = @('noreply@facultydomain.com') #To email the entire student body.
}

#GAM To create user before GADS runs so that when the password is set it uses GAPS to set it in G Suite.
$GAMprecreateUser = $True

#GAM Default Organizational Unit to put new account.
$GAMDefaultOrg = "/Students"

#If GAM runs into a duplicate while trying to precreate do you want to stop so you can sort it out before the account is taken over by the new AD account?
$GAMStopOnDuplicate = $True

#Do you want GAM to set the moderation rules on the groups at creation time?  This affects the students groups for building and class as well.
$GAMsetModerationRules = $True

#Do you want GAM to verify that moderation rules are set for ALL groups under OU=Students?
$GAMVerifyModerationRules = $True

#Do you want Homerooms in eSchool to be built into the schedule as a period 0 for Clever?
$homeroomsScheduled = $False

#Do you want Activities in eSchool to be built into the schedule as a period 99 for Clever?
$activitiesScheduled = $False
$activitiesBuildings = @(13,15,16,703) #Array for which buildings you want to have activity schedules built for
$activitiesInclude = @("*") #Array of strings. Remove the astrick if you want to only include specific activities to use in the schedule.
$activitiesIgnore = @("*Volleyball*","*Golf*","*Football*") #Array of strings to ignore and not build a schedule for.

#Do you want this script to be disabled via a remote server in case of upgrades to eSchool that might have unintended consequences?
#This will also check the version and tell you how to upgrade.
$remoteCheck = $False

#You can host your own response server and have control over this. Otherwise this one is provided and controlled by Craig Millsap.
$remoteCheckURL = "https://www.camtechcs.com/automated_students/statuscheck.php"

#Seriously you better know what you're doing or bad things will happen. I definitely believe you should have both Domain Admin and run the script as Administrator.
#$DoNotRequireDomainAdmin = $True
#$DoNotRequireAdminstrator = $True

#Skip disabling accounts that have a EmployeeNumber but are not in the exclusions.csv or the students table.
$SkipDisablingAccounts = $True

#How many days do you want to let a student still have access to their account before disabling/suspending it?
$RetainStudentsDays = 0

# Clever reports are based on a students full schedule. If you put in a drop date and enroll them in a different full year class the enrollments.csv from
# Cognos will not show the dropped class. In 2020 the changes to schedules from Virtual to Onsite started almost 40 days early. This caused students to be unenrolled
# from their classes in Clever. This is a work around to maintain the existing enrollments for X days after it is no longer reported in the enrollments.csv from Cognos.
$RetainEnrollmentsDays = 0
$RetainEnrollmentBuildings = @(703,15)

#If you don't specify and you use the script in the PostProcessingScripts folder you will get all buildings.
$TeacherPasswordsForBuildings = '13,16,703' 

#This will permanently delete any accounts that have a sign in status of NEVER. It will also permanently delete accounts that haven't signed in the last 5 years unless overridden with $GWorkspaceAfterxMonths
$GWorkspaceDeleteSuspendedAccounts = $False
$GWorkspaceDeleteAfterXMonths = 60 #defaults to 5 years if you don't specify.

#These are for testing purposes. Leave them commented out when running in production.
#$DisablePostProcessingScript = $True
#$SkipDownloadingReports = $True
#$SkipTableCleanup = $True
$Staging = $True