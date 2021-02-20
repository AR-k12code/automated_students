###
#
# Name: Student Import Script
# Created by: Craig Millsap, Gentry Public Schools, cmillsap@facultydomain.com
# Date: 7/30/2020
#
# All data is now stored in an sqlite database
# Empty brackets are so I can collapse a section in the editor.

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)

Import-Module PSSQLite
. $currentPath\settings.ps1

$logfile = "$currentPath\logs\student-export-files-$(get-date -f yyyy-MM-dd-HH-mm-ss).log"

if (-Not(Test-Path "$currentPath\logs")) { New-Item -ItemType Directory "$currentPath\logs" -Force }
try { Start-Transcript("$logfile") } catch { Stop-TranScript; Start-Transcript("$logfile") }

if (-Not(Test-Path "$currentPath\exports")) { New-Item -ItemType Directory "$currentPath\exports" -Force }

$year = [int](Get-Date -Format yyyy)
if ([int](Get-Date -Format MM) -le 6) {
    $year = $year - 1
}

#############
# Functions #
#############

. $currentPath\z_functions.ps1

######################
# Build Export Files #
######################

############
# HMH HMO  #
############
if ($False) {
    if (-Not(Test-Path $currentPath\exports\HMH)) { New-Item -ItemType Directory -Path "$currentPath\exports\HMH" -Force }
    $q = "SELECT
        $year AS 'SCHOOLYEAR',
        'S' AS 'ROLE',
        Student_id AS 'LASID',
        '' AS 'SASID',
        students.First_name AS 'FIRSTNAME',
        '' AS 'MIDDLENAME',
        students.Last_name AS 'LASTNAME',
        Grade AS 'GRADE',
        students.Student_email AS 'USERNAME',
        '' AS 'PASSWORD',
        'MDR' AS 'ORGANIZATIONTYPEID',
        CASE School_id
            WHEN 16 THEN '00020787'
            WHEN 13 THEN '04912637'
            WHEN 15 THEN '02111485'
            WHEN 703 THEN '00020799'
        END AS 'ORGANIZATIONID',
        students.Student_email AS 'PRIMARYEMAIL',
        '' AS 'HMHAPPLICATIONS'
        FROM students"
    $hmhusers = Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' # | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\HMH\users.csv -Force

    $q = "SELECT
        $year AS 'SCHOOLYEAR',
        'T' AS 'ROLE',
        Teacher_id AS 'LASID',
        '' AS 'SASID',
        teachers.First_name AS 'FIRSTNAME',
        '' AS 'MIDDLENAME',
        teachers.Last_name AS 'LASTNAME',
        CASE School_id
            WHEN 16 THEN 'K-2'
            WHEN 13 THEN '3-5'
            WHEN 15 THEN '6-8'
            WHEN 703 THEN '9-12'
        END
            AS 'Grade',
        teachers.Teacher_email AS 'USERNAME',
        '' AS 'PASSWORD',
        'MDR' AS 'ORGANIZATIONTYPEID',
        CASE School_id
            WHEN 16 THEN '00020787'
            WHEN 13 THEN '04912637'
            WHEN 15 THEN '02111485'
            WHEN 703 THEN '00020799'
        END
            AS 'ORGANIZATIONID',
        teachers.Teacher_email AS 'PRIMARYEMAIL',
        '' AS 'HMHAPPLICATIONS'
        FROM teachers
        WHERE teachers.Teacher_email != ''"
    $hmhusers += Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP'
    $hmhusers | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\HMH\users.csv -Force

    $teachers = Invoke-SqliteQuery -database $database -Query "select * from teachers WHERE Teacher_email != ''"
    #HMH Teachers Classes and Assignments

    # $header = "SCHOOLYEAR,CLASSLOCALID,COURSEID,COURSENAME,COURSESUBJECT,CLASSNAME,CLASSDESCRIPTION,CLASSPERIOD,ORGANIZATIONTYPEID,ORGANIZATIONID,GRADE,TERMID,HMHAPPLICATIONS"
    # Out-File -Encoding ASCII -filepath $currentPath\exports\HMH\class.csv -inputobject $header -Force
    # $hmhclass = Invoke-SqliteQuery -DataSource $database -Query '/*
    # SCHOOLYEAR,CLASSLOCALID,COURSEID,COURSENAME,COURSESUBJECT,CLASSNAME,CLASSDESCRIPTION,CLASSPERIOD,ORGANIZATIONTYPEID,ORGANIZATIONID,GRADE,TERMID,HMHAPPLICATIONS /*
    # SELECT "" AS "SCHOOLYEAR","" AS "CLASSLOCALID","" AS "COURSEID","" AS "COURSENAME","" AS "COURSESUBJECT","" AS "CLASSNAME","" AS "CLASSDESCRIPTION","" AS "CLASSPERIOD","" AS "ORGANIZATIONTYPEID","" AS "ORGANIZATIONID","" AS "GRADE","" AS "TERMID","" AS "HMHAPPLICATIONS"'
    #empty object with the right property names for appending to later.
    $hmhclassassignments = @()
    #Invoke-SqliteQuery -DataSource $database -Query '/* SCHOOLYEAR,CLASSLOCALID,LASID,ROLE,POSITION */
    #SELECT "" AS "SCHOOLYEAR","" AS "CLASSLOCALID","" AS "LASID", "" AS "ROLE", "" AS "POSITION"' #This builds our object to append to.

    #HS Teachers are for their curiculum not for remediation.
    $hmhteachers = @('wpipkin@facultydomain.com','lmoore@facultydomain.com','jmadding@facultydomain.com','ejones@facultydomain.com','jpierce@facultydomain.com','tswicegood@facultydomain.com') #HS
    $hmhteachers += $teachers | Where-Object { $PSItem.'School_id' -eq 16 } | Select-Object -ExpandProperty Teacher_email #Primary
    $hmhteachers += $teachers | Where-Object { $PSItem.'School_id' -eq 13 } | Select-Object -ExpandProperty Teacher_email #Intermediate

    $list = "`(""$($hmhteachers -join '","')""`)"
    $hmhclasses = Invoke-SqliteQuery -DataSource $database -Query "/*
    SCHOOLYEAR,CLASSLOCALID,COURSEID,COURSENAME,COURSESUBJECT,CLASSNAME,CLASSDESCRIPTION,CLASSPERIOD,ORGANIZATIONTYPEID,ORGANIZATIONID,GRADE,TERMID,HMHAPPLICATIONS
    */
    SELECT
        $year AS 'SCHOOLYEAR',
        '' AS 'CLASSLOCALID',
        schedules.Course_number AS 'COURSEID',
        '' AS 'COURSENAME',
        '' AS 'COURSESUBJECT',
        '' AS 'CLASSNAME',
        sections_grouped.Name AS 'CLASSDESCRIPTION',
        schedules.Period AS 'CLASSPERIOD',
        'MDR' AS 'ORGANIZATIONTYPEID',
        CASE schedules.School_id
            WHEN 16 THEN '00020787'
            WHEN 13 THEN '04912637'
            WHEN 15 THEN '02111485'
            WHEN 703 THEN '00020799'
        END
            AS 'ORGANIZATIONID',
        CASE schedules.School_id
            WHEN 16 THEN '2'
            WHEN 13 THEN '5'
            WHEN 15 THEN '8'
            WHEN 703 THEN '12'
        END
            AS 'GRADE',
        '' AS 'TERMID',
        '' AS 'HMHAPPLICATIONS',
        schedules.School_id,
        schedules.Section_id,
        schedules.Teacher_id,
        schedules.Section_number,
        schedules.Teacher_email
        FROM schedules
        LEFT JOIN sections_grouped ON schedules.Section_id = sections_grouped.Section_id
        AND Schedules.School_id = sections_grouped.School_id
        WHERE Teacher_email IN $list
        GROUP BY schedules.Section_id"
    #"$year,$classid,$courseid,,,$classname,$classdescription,$classperiod,MDR,$hmhorgid,$grade,,"

    $hmhclasses | ForEach-Object {

        switch ($PSItem.'School_id') {
            16 { $buildingShortName = "GPS"; }
            13 { $buildingShortName = "GIS"; }
            15 { $buildingShortName = "GMS"; }
            703 { $buildingShortName = "GHS"; }
            default { return }
        }

        $teacherid = $PSItem.'Teacher_id'
        $classlocalid = $buildingShortName + '-' + $($PSItem.'Teacher_email').Split('@')[0] + "-" + $PSItem.'CLASSPERIOD' + "-" + $(Remove-Spaces(Remove-SpecialCharacters(Remove-Dashes($PSItem.'CLASSDESCRIPTION')))) + "-" + $PSItem.'Section_number'
        
        #write-host $classlocalid
        
        #generate the name for the course id.
        $PSItem.'CLASSLOCALID' = $classlocalid
        $PSItem.'COURSENAME' = $classlocalid
        
        # $classline = "$year,$classid,$courseid,,,$classname,$classdescription,$classperiod,MDR,$hmhorgid,$grade,,"
        # Out-File -Encoding ASCII -filepath $currentPath\exports\HMH\class.csv -inputobject $classline -Append

        #Assign teacher
        $hmhclassassignments += Invoke-SqliteQuery -DataSource $database -Query "/* SCHOOLYEAR,CLASSLOCALID,LASID,ROLE,POSITION */
            SELECT $year AS 'SCHOOLYEAR',
            ""$classlocalid"" AS 'CLASSLOCALID',
            $teacherid AS 'LASID',
            'T' AS 'ROLE',
            '' AS 'POSITION'"

        #Assign students based on the section_id.
        $hmhclassassignments += Invoke-SqliteQuery -DataSource $database -Query "/* SCHOOLYEAR,CLASSLOCALID,LASID,ROLE,POSITION */
            SELECT $year AS 'SCHOOLYEAR',
            ""$classlocalid"" AS 'CLASSLOCALID',
            Student_id AS 'LASID',
            'S' AS 'ROLE',
            '' AS 'POSITION'
            FROM schedules
            WHERE Section_id = $($PSItem.'Section_id')"
            
    }

    $hmhclasses | Select-Object -Property * -ExcludeProperty School_id,Section_id,Teacher_id,Section_number,Teacher_email | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\HMH\class.csv -Force
    $hmhclassassignments | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\HMH\classassignments.csv -Force

    #HMO Must Be Zipped
    Compress-Archive -LiteralPath $currentPath\exports\HMH\users.csv,$currentPath\exports\HMH\class.csv,$currentPath\exports\HMH\classassignments.csv -CompressionLevel Optimal -DestinationPath $currentPath\exports\HMH\import.zip -Force

    $server = 'sftp.idma.app.hmhco.com'
    $username = ''
    $password = ''

    #10/2020 DO NOT TURN ON UNTIL YOU FIX THE Course_description vs Course_name mishap where you borked it. Match existing until next school year.
    #Start-Process -FilePath "$currentPath\bin\pscp.exe" -ArgumentList "-pw $($password) $currentPath\files\HMH\import.zip $($username)@$($server):" -PassThru -Wait -NoNewWindow
}

###############
# HMH SAM WEB #
###############
if ($True) {
    if (-Not(Test-Path $currentPath\exports\HMH)) { New-Item -ItemType Directory -Path "$currentPath\exports\HMH" -Force }

    $hmhteachers = @('kcarr@facultydomain.com','wpipkin@facultydomain.com','mmontgomery@facultydomain.com','dbuss@facultydomain.com') #HS
    $hmhteachers += @('rsummers@facultydomain.com','smccollum@facultydomain.com','sellison@facultydomain.com','tsweeten@facultydomain.com','scunningham@facultydomain.com') #MS
    $hmhteachers += @('kmadding@facultydomain.com') #IS
    $hmhteachers += @('bcordeiro@facultydomain.com','jellis@facultydomain.com','serks@facultydomain.com','cking@facultydomain.com','mwilson@facultydomain.com','tmcnelly@facultydomain.com') #Something to do with Primary School.

    $hmhstudentschedules = @()

    $list = "`(""$($hmhteachers -join '","')""`)"
    $q = "SELECT
        teachers.Teacher_id AS 'DISTRICT_USER_ID',
		teachers.Teacher_id AS 'SPS_ID',
		'' AS PREFIX,
		teachers.First_name AS 'FIRST_NAME',
		teachers.Last_name AS 'LAST_NAME',
		'' AS 'TITLE',
		'' AS 'SUFFIX',
		teachers.Teacher_email AS 'EMAIL',
		REPLACE(teachers.Teacher_email,'@facultydomain.com','') AS 'USER_NAME',
		'' AS 'PASSWORD',
		CASE schedules.School_id
            WHEN 16 THEN 'Gentry Primary School'
            WHEN 13 THEN 'Gentry Intermediate School'
            WHEN 15 THEN 'Gentry Middle School'
            WHEN 703 THEN 'Gentry High School'
        END AS 'SCHOOL_NAME',
        CASE schedules.School_id
            WHEN 16 THEN 'GPS'
            WHEN 13 THEN 'GIS'
            WHEN 15 THEN 'GMS'
            WHEN 703 THEN 'GHS'
        END AS 'SCHOOL_SHORT_NAME',
		'' AS CLASS_NAME,
		teachers.Teacher_email AS 'EXTERNAL_ID',
		'Y' AS 'LAST_COL',
		schedules.Section_id,
        /* sections_grouped.Course_description, #FIX THIS LATER */
        sections_grouped.Name,
        schedules.Period,
        schedules.Section_number
		FROM schedules
        LEFT JOIN sections_grouped ON schedules.Section_id = sections_grouped.Section_id
		LEFT JOIN teachers ON schedules.Teacher_id = teachers.Teacher_id
        AND Schedules.School_id = sections_grouped.School_id
        WHERE teachers.Teacher_email IN $list
        AND schedules.Period != 'ENC'
        GROUP BY schedules.Section_id
        ORDER BY EMAIL, schedules.Period"

    $hmhschedules = Invoke-SqliteQuery -Database $database -Query $q

    #we need to build the CLASS_NAME then enroll students per the sections.
    $hmhschedules | ForEach-Object {
        #Class Names should be BUILDING-USER-PERIOD-COURSENAME-SECTION
        $classname = $PSItem.'SCHOOL_SHORT_NAME' + '-' + $PSItem.'USER_NAME' + '-' + $PSItem.'Period' + '-' + $(Remove-Spaces(Remove-SpecialCharacters(Remove-Dashes($PSItem.'Name')))) + '-' + $PSItem.'Section_number'
        $PSitem.'CLASS_NAME' = $classname
        $sectionid = $PSItem.'Section_id'
        #write-host $classname

        #pull students schedules and enroll them in this class.
        $q = "select
            REPLACE(students.Student_email,'@studentdomain.com','') AS USER_NAME,
            '' AS PASSWORD,
            students.Student_id AS SIS_ID,
            students.First_name AS FIRST_NAME,
            '' AS MIDDLE_NAME,
            students.Last_name AS LAST_NAME,
            students.Grade AS GRADE,
            CASE schedules.School_id
                WHEN 16 THEN 'Gentry Primary School'
                WHEN 13 THEN 'Gentry Intermediate School'
                WHEN 15 THEN 'Gentry Middle School'
                WHEN 703 THEN 'Gentry High School'
            END AS 'SCHOOL_NAME',
            '' AS CLASS_NAME,
            '' AS LEXILE_SCORE,
            '' AS LEXILE_MOD_DATE,
            '' AS ETHNIC_CAUCASIAN,
            '' AS ETHNIC_AFRICAN_AM,
            '' AS ETHNIC_HISPANIC,
            '' AS ETHNIC_PACIFIC_ISL,
            '' AS ETHNIC_AM_IND_AK_NATIVE,
            '' AS ETHNIC_ASIAN,
            '' AS ETHNIC_TWO_OR_MORE_RACES,
            '' AS GENDER_MALE,
            '' AS GENDER_FEMALE,
            '' AS AYP_ECON_DISADVANTAGED,
            '' AS AYP_LTD_ENGLISH_PROFICIENCY,
            '' AS AYP_GIFTED_TALENTED,
            '' AS AYP_MIGRANT,
            '' AS AYP_WITH_DISABILITIES,
            students.Student_email AS EXTERNAL_ID,
            'Y' AS LAST_COL
            FROM schedules
            INNER JOIN students ON students.Student_id = schedules.Student_id
            WHERE schedules.Section_id = $sectionid"

        $hmhstudents = Invoke-SqliteQuery -Database $database -Query $q
        
        $hmhstudents | ForEach-Object {
            $PSItem.'CLASS_NAME' = $classname
        }

        $hmhstudentschedules += $hmhstudents
    }

    #teachers
    $hmhschedules | Select-Object -Property * -ExcludeProperty Section_id,Name,Period,Section_number,SCHOOL_SHORT_NAME | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\hmh\HMH-SAM-Teachers.csv -Force

    #students
    $hmhstudentschedules | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\hmh\HMH-SAM-Students.csv -Force

}
####################################################
# SwiftK12 Export
####################################################
if ($True) {
    $q = 'SELECT 
        students.School_id,students.Student_id,students.First_name,students.Last_name,students.Grade,
        contacts.Contact_priority,contacts.Contact_firstname,contacts.Contact_lastname,contacts.Contact_phonetype,contacts.Contact_phonenumber,contacts.Contact_email,
        students.Home_language,
        students_extras.Student_houseteam,
        students_extras.Student_homeroom,
        teachers.Teacher_email AS Student_hmrmteacheremail,
        transportation.Student_BusNumFrom,
        transportation.Student_BusNumTo
        FROM students
        LEFT JOIN contacts ON students.Student_id = contacts.Student_id
        LEFT JOIN students_extras ON students.Student_id = students_extras.Student_id
        LEFT JOIN teachers ON students_extras.Student_hrmtid = teachers.Teacher_id
        LEFT JOIN transportation ON students.Student_id = transportation.Student_id
        ORDER BY students.Student_id'
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\SwiftK12.csv -Force
    Copy-Item $currentPath\exports\SwiftK12.csv $currentPath\..\swiftk12\files\SwiftK12_Students.csv -Force
}

####################################################
# Destiny
####################################################
if ($True) {
    $q = '/*
    SiteShortName,Barcode,DistrictID,LastName,FirstName,PatronType,AccessLevel,Status,Gender,Homeroom,GradeLevel,GraduationYear,UserName,Password,EmailPrimary
    */

    SELECT
        students.School_id AS SiteShortName,
        students.Student_id AS Barcode,
        students.Student_id AS DistrictID,
        students.Last_name as LastName,
        students.First_name as FirstName,
        "Student" AS PatronType,
        "Patron" as AccessLevel,
        "A" as Status,
        students.Gender as Gender,
        lower(substr(teachers.First_name,1,1) || teachers.Last_name) as Homeroom,
        REPLACE(students.Grade,"Kindergarten","K") as GradeLevel,
        students_extras.Student_gradyr as GraduationYear,
        REPLACE(students.Student_email,"@studentdomain.com","") as Username,
        "8icriCraEvfERSAcrEtAba" as Password,
        students.Student_email as EmailPrimary
        FROM students
        LEFT JOIN students_extras 
        ON students.Student_id = students_extras.Student_id
        LEFT JOIN teachers
        on teachers.Teacher_id = students_extras.Student_hrmtid'
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' | Select-Object -Property * -ExcludeProperty Timestamp | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\destiny.csv -Force
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c C:\Scripts\destiny\process_students.bat" -NoNewWindow -Wait
}
    
####################################################
# IDVille
####################################################
if ($True) {
    if (-Not(Test-Path $currentPath\exports\idville)) { New-Item -ItemType Directory -Path "$currentPath\exports\idville" -Force }
    $q = '/*
    Proper First,Proper Last,Grade,Graduation Year,Picture Path,Student ID,QRCode Path
    $picpath = "g:\Shared drives\LifeTouch\combined\current\images\$($id).jpg"
    $qrpath = "g:\Shared drives\LifeTouch\Combined\QR_Codes\$($id).png"
    */
    SELECT
        students.School_id,
        students.First_name AS "Proper First",
        students.Last_name AS "Proper Last",
        REPLACE(students.Grade, "Kindergarten", "K") AS Grade,
        students_extras.Student_gradyr as "Graduation Year",
        ("g:\Shared drives\LifeTouch\combined\current\images\" || students.Student_id || ".jpg") as "Picture Path",
        ("g:\Shared drives\LifeTouch\Combined\QR_Codes\" || students.Student_id || ".jpg") as "QRCode Path"
        FROM students
        INNER JOIN students_extras
        ON students.Student_id = students_extras.Student_id
        ORDER BY Grade'
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' | Where-Object { $PSItem.'School_id' -eq 703 } | Select-Object -Property * -ExcludeProperty School_id | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\idville\GHSCC.csv -Force
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' | Where-Object { $PSItem.'School_id' -eq 15 } | Select-Object -Property * -ExcludeProperty School_id | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\idville\GMS.csv -Force
    #IDVille Upload
    gam user technology@facultydomain.com update drivefile id 1PW3eViDJ39t2oir2Y1I0000000000000 localfile "$currentPath\exports\idville\GHSCC.csv" parentid 1NMuFbiwg7rJRhuPnFWfP000000000000
    gam user technology@facultydomain.com update drivefile id 1BsXbAvlEBYsMpx4zJ0B0000000000000 localfile "$currentPath\exports\idville\GMS.csv" parentid 1xielnVxbn_BWY12oNc00000000000000
}

####################################################
# Apex
####################################################
if ($True) {
    $q ='/*
    RecordNumber,RecordAction,Role,ImportOrgID,OrgName,ImportUserID,FirstName,LastName,Email,LoginID,LoginPW
    */
    SELECT
        "" AS RecordNumber,
        "A" AS RecordAction,
        "S" AS Role,
        students.School_id AS ImportOrgID,
        REPLACE(REPLACE(REPLACE(REPLACE(School_id, 15, "GMS"), 703, "GHSCC"), 16, "GPS"), 13, "GIS") AS OrgName,
        students.Student_id AS ImportUserID,
        students.First_name AS FirstName,
        students.Last_name AS LastName,
        students.Student_email as Email,
        students.Student_email as LoginID,
        passwords.Student_password AS LoginPW
        FROM students
        LEFT JOIN passwords
        ON students.Student_id = passwords.Student_id
        WHERE students.School_id IN (703,15)
        ORDER BY School_id,students.Student_id'
    $apex = Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP'
    $apex | ForEach-Object { $apexcount++; $PSItem.'RecordNumber' = $apexcount }
    $apex | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\apex.csv -Force
}

###############
# Square
###############
if ($True) {
    $q = 'SELECT
    students.First_name AS "First Name",
    students.Last_name AS "Last Name",
    "" AS "Company Name",
    students.Student_email AS "Email Address",
    "" AS "Phone Number",
    "" AS "Street Address 1",
    "" AS "Street Address 2",
    "" AS "City",
    "" AS "State",
    "" AS "Postal Code",
    students.Student_id AS "Reference ID",
    "" as Birthday
    FROM students
    WHERE students.School_id IN (703,15)
    UNION
    SELECT
    teachers.First_name AS "First Name",
    teachers.Last_name AS "Last Name",
    "" AS "Company Name",
    teachers.Teacher_email AS "Email Address",
    "" AS "Phone Number",
    "" AS "Street Address 1",
    "" AS "Street Address 2",
    "" AS "City",
    "" AS "State",
    "" AS "Postal Code",
    teachers.Teacher_id AS "Reference ID",
    "" as Birthday
    FROM teachers'
    Invoke-SqliteQuery -Database $database -Query $q | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\squareup.csv -Force
}

############
# Student Info for Badge System  #
############
if ($True) {
    $q = "SELECT 
    students.Student_id AS 'Student ID',
    First_name AS Firstname,
    Last_name AS Lastname,
    students_extras.Student_gradyr AS 'Graduation Year',
    students.Grade AS Grade,
    students.Gender AS Gender,
    students.School_id AS 'Current Building'
    FROM
    students
    LEFT JOIN students_extras ON students.Student_id = students_extras.Student_id
    ORDER BY 'Current Building',Grade,Lastname"
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\badges.csv -Force
    Copy-Item "$currentPath\exports\badges.csv" "\\badge-pc\c$\data\badges.csv" -Force
}

###################
# NutriKids Mosaic
###################
if ($True) {
    $q = "SELECT
    students.First_name AS 'First Name',
    students.Last_name AS 'Last Name',
    students.Student_id AS 'Student ID',
    students.State_id AS 'State ID #',
    students.School_id AS 'School Code',
    CASE students.Grade
        WHEN 'Kindergarten' THEN 'KF'
        WHEN 'Prekindergarten' THEN 'PK'
        ELSE students.Grade
    END AS 'Grade',
    students.DOB AS 'Student DOB',
    students_extras.Student_mealstatus AS 'Meal Status',
    /* (SELECT Contact_firstname FROM contacts WHERE contacts.Student_id = students.Student_id ORDER BY Contact_priority LIMIT 1) AS 'Guardian First Name',
    (SELECT Contact_Lastname FROM contacts WHERE contacts.Student_id = students.Student_id ORDER BY Contact_priority LIMIT 1) AS 'Guardian Last Name', */
    substr(Contact_name,1,INSTR(Contact_name,' ')) AS 'Guardian First Name',
	substr(Contact_name,INSTR(Contact_name,' '),LENGTH(Contact_name)) AS 'Guardian Last Name',
    students.Contact_email AS 'Guardian Email',
    students.Contact_phone AS 'Guardian Phone',
    students.Student_street AS 'Mailing Address',
    students.Student_city AS 'City',
    students.Student_state AS 'State',
    students.Student_zip AS 'Zip',
    students.Student_email AS 'Student Email',
    students_extras.Student_gradyr AS 'Year of Graduation',
    1 AS 'Active',
    (teachers.First_name || ', ' || teachers.Last_name) AS 'Homeroom Teacher'
    from students
    LEFT JOIN
    students_extras ON students.Student_id = students_extras.Student_id
    LEFT JOIN
    teachers ON students_extras.Student_hrmtid = teachers.Teacher_id"
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File "$currentPath\exports\NutriKids Mosaic.csv" -Force
    Start-Process -FilePath "$currentPath\bin\pscp.exe" -ArgumentList "-pw XXXXXXXX ""$currentPath\exports\NutriKids Mosaic.csv"" username@Cpsftp.heartlandmosaic.com:" -PassThru -Wait -NoNewWindow
}

####################################################
# Destiny2
####################################################
if ($True) {
    $q = 'SELECT
    students.Student_id AS "Student ID",
    students.Last_name as LastName,
    students.First_name as FirstName,
    students.Gender as Gender,
    /* lower(substr(teachers.First_name,1,1) || teachers.Last_name) as Homeroom, */
    (students_extras.Student_homeroom || " - " || teachers.Last_name) as Homeroom,
    CASE students.Grade
        WHEN "Kindergarten" THEN "KF"
        WHEN "PreKindergarten" THEN "PK"
        ELSE students.Grade
    END AS Grade,
    students_extras.Student_gradyr as GraduationYear,
    students.DOB AS Birthdate,
    students.Student_email as Email,
    students.School_id AS Building
    FROM students
    LEFT JOIN students_extras 
    ON students.Student_id = students_extras.Student_id
    LEFT JOIN teachers
    on teachers.Teacher_id = students_extras.Student_hrmtid
    ORDER BY Building,Grade'
    Invoke-SqliteQuery -Database $database -Query $q -ErrorAction 'STOP' |ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $currentPath\exports\destiny2.csv -Force
    Start-Process -FilePath "$currentPath\bin\pscp.exe" -ArgumentList "-pw XXXXXXX $currentPath\exports\destiny2.csv username@data.follettsoftware.com:""/patrons/Destiny Student File.csv""" -PassThru -Wait -NoNewWindow
}

###############################
# SEAS - Cognos Report Download
###############################
Invoke-Expression "$($currentPath)\..\CognosDownload.ps1 -report ""SEAS Student File"" -RunReport -savepath ""$currentPath\exports\"" -username $eSchoolUsername -espdsn $eSchooldsn -extension csv -reportwait $reportWait -ReportStudio"
Start-Process -FilePath "$currentPath\bin\pscp.exe" -ArgumentList "-pw XXXXXXXXX ""$currentPath\exports\SEAS Student File.csv"" USERNAME@ftp.seasweb.net:" -PassThru -Wait -NoNewWindow


##############################
# Upload to Google Drive
##############################

Start-Process -FilePath "rclone.exe" -ArgumentList "--config c:\scripts\rclone\rclone.conf -v sync $currentPath\exports google-drive:ImportFiles/" -NoNewWindow -Wait

Stop-Transcript