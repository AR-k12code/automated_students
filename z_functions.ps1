#############
# Functions #
#############

function Convert-ToProperCase ([String]$in) {
    $in = $in.Tolower()
    $textInfo = [System.Threading.Thread]::CurrentThread.CurrentCulture.TextInfo  
    return $textInfo.ToTitleCase($in)
}

function Remove-SpecialCharacters ([String]$in) {
    #Its faster to put this on one line but readability
    $in = $in -replace("\(","") #Remove ('s
    $in = $in -replace("\)","") #Remove )'s
    $in = $in -replace("\.","") #Remove Periods
    $in = $in -replace("\'","") #Remove Apostrophies
    $in = $in -replace('\*','') #Remove Asterisks
    $in = $in -replace(',','') #Remove Commas
    $in = $in -replace('\\','') #Remove Back Slash
    $in = $in -replace('/','') #Remove Foward Slash
    $in = $in -replace("'",'') #Remove Single Quote
    $in = $in -replace('"','') #Remove Double Quote
    return $in
}

function Remove-Spaces ([String]$in) {
 $in = $in -replace(' ','') #Remove Spaces
 return $in
}

function Remove-Dashes ([String]$in) {
    $in = $in -replace('-','') #Remove Spaces
    return $in
   }

function Get-RandomCharacters($length, $characters) { 
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    return [String]$characters[$random]
}

function New-HomeDirectory ([string]$username, [string]$path) {
    if (Test-Path $path) {
        write-host "Warning: Home directory already exists for user $username at $path" -ForegroundColor Yellow
    } else {
        Write-host "Info: Creating new home directory for $username at $path." -ForegroundColor Yellow
        $createFolder = New-Item $path -type Directory -Force
    }
    
    Write-host "Info: Ensure correct permissions on $path for $username." -ForegroundColor Yellow
    Clear-NTFSAccess -Path $path
    Enable-NTFSAccessInheritance -Path $path
    
    #Something here doesn't work. It has to be a timing issue with the account not existing on all servers yet.
    #we get an error about not finding the account. Lets try waiting...
    do {
        Start-Sleep -Seconds 1;
        $numErrors++
        if ($numErrors -gt 10) { return $False }
        #write-host "Checking that account $username exists to apply home directory permissions"
    } until ($account = Get-ADUser $username)

    Add-NTFSAccess -Path $path -Account $account.SID -AccessRights Modify
}

function Reset-HomeDirPermissions ([string]$username, [string]$path) {
    #This function also creates the folder.
    if (-Not(Test-Path $path)) {
        #Folder not found. Test next Folder Up.
        write-host  "Warning: Home Directory folder does not exist. Attempting to create."
        $rootPath = $($($homedir).split('\')[0..($($homedir.split('\')).Count -2 )] -join '\')
        if (Test-Path $rootPath) {
            New-Item $path -ItemType Directory
        } else {
            Write-Host "Error: $rootPath is either missing or unaccessible. Please fix before running again." -ForegroundColor Red
        }
    }
    Clear-NTFSAccess -Path $path
    Enable-NTFSAccessInheritance -Path $path
    
    #This shouldn't be needed anymore.
    #do { start-sleep 3; write-host "Checking that account $username exists to apply home directory permissions" } until ($(get-aduser $username).samaccountname -eq "$username")
    
    Add-NTFSAccess -Path $path -Account $username -AccessRights Modify
}

function Grant-PasswordResetOnOU ($group, $ou) {
    $adgroup = Get-ADGroup $group
    $acl = Get-ACL "AD:\$($ou)"
    $grpSID = New-Object System.Security.Principal.SecurityIdentifier ($adgroup).SID
    $acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $grpSID,"ExtendedRight","Allow",([GUID]("00299570-246d-11d0-a768-00aa006e0529")).guid,"Descendents",([GUID]("bf967aba-0de6-11d0-a285-00aa003049e2")).guid))
    $acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $grpSID,"WriteProperty","Allow",([GUID]("bf967a0a-0de6-11d0-a285-00aa003049e2")).guid,"Descendents",([GUID]("bf967aba-0de6-11d0-a285-00aa003049e2")).guid))
    $acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $grpSID,"ReadProperty","Allow",([GUID]("bf967a0a-0de6-11d0-a285-00aa003049e2")).guid,"Descendents",([GUID]("bf967aba-0de6-11d0-a285-00aa003049e2")).guid))
    #Apply Additional ACL rules above
    Set-ACL -Path "AD:\$($ou)" -AclObject $acl
}

Function Get-NextAvailableUsername ($username, $principalName, $homeDirGradYR, $firstName, $lastName, $middleInitial) {
    #we already know there is a conflict.
    #we need another function to generate usernames so we can pass it to that as many times as needed.
    $firstName = Remove-SpecialCharacters(Remove-Spaces($firstName))
    $lastName = Remove-SpecialCharacters(Remove-Spaces($lastName))

    $newUser = @{}
    $principalDomain = $principalName.split('@')[1]
    
    if ($middleInitial.length -eq 1) {
        #first lets try adding the middle initial. #this is going to have to be done with a replace.
        $newUser.'username' = $username -replace "$firstName","$($firstName)$($middleInitial)"
        #i need to truncate the lastname from the username one character if its longer than 20 with the added middleInitial.
        #this must be done in reverse to find match. We might have already truncated the last name.
        #$($newUser.'username').Length
        if ($($newUser.'username').Length -gt 20) {
            for ($i = $lastName.Length; $i -ge 0; $i--) {
                $lastNameCheck = $lastName.Substring(0,$i)
                #write-host "searching for $lastNameCheck"
                if ($newUser.'username' | Select-String -Pattern "$lastNameCheck") {
                    #replace and break
                    #write-host "found match for $lastNameCheck"
                    $newUser.'username' = $newUser.'username' -replace "$lastNameCheck",$lastNameCheck.Substring(0,$i-1)
                    #write-host "It is now $($newUser.'username')"
                    break
                }
            }
        }
        
        #Student Template 7 & 8 do not use the full name and will always fit inside the 20 characters. This means we can use the generated username for the userprincipal as it will ALWAYS match.
        if (@(7,8) -contains $stuTemplate) {
            $newUser.'principalName' = $($newUser.'username') + $principalDomain
        } else {
            $newUser.'principalName' = $principalName -replace "$firstName","$($firstName)$($middleInitial)"
        }

        $newUser.'homeDir' = "$($homeDirGradYR)\$($newUser.'username')"
        #write-host "We should return $($newUser.'username') back to the main script."
        #now test and return
        #if ($(Get-AdUser -Filter "(SamAccountName -eq ""$($newUser.'username')"")") -or $(Get-AdUser -Filter "(UserPrincipalName -eq ""$($newUser.'principalName')"")")) {
        #if ($(Get-AdUser -Filter "(SamAccountName -eq ""$($newUser.'username')"") -or (UserPrincipalName -eq ""$($newUser.'principalName')"")")) {
        if ($(Get-AdUser -Filter "(ObjectGUID -ne ""$existingAccountGUID"") -and ((SamAccountName -eq ""$($newUser.'username')"") -or (UserPrincipalName -eq ""$($newUser.'principalName')""))")) {
            Write-Host "$($newUser.'username'),$($newUser.'PrincipalName') is also not available."
        } else {
            return $newUser
        }
    }

    #i need to truncate the lastname from the username one character if its longer than 20 with the added middleInitial.
    #this must be done in reverse to find match. We might have already truncated the last name.
    
    $errors = 0
    $newUser.'username' = $username
    for ($i = 2; $i -le 10; $i++) { #If we have more than 10 conflicts we need to rethink our approach.
        for ($j = $lastName.Length; $j -ge 0; $j--) {
            $lastNameCheck = $($lastName).Substring(0,$j)
            if ($username | Select-String -Pattern "$lastNameCheck") {
                #replace and break
                $newLastName = $lastNameCheck.Substring(0,$j) + [string]$i
                $newUser.'username' = $username -replace "$lastNameCheck","$newLastName"
                break
            }
        }

        #what if adding the 2 to the end of the name now makes it too long?
        if ($($newUser.'username').Length -gt 20) {
            for ($k = $lastName.Length; $k -ge 0; $k--) {
                $lastNameCheck = $lastName.Substring(0,$k)
                #write-host "searching for $lastNameCheck"
                if ($newUser.'username' | Select-String -Pattern "$lastNameCheck") {
                    #replace and break
                    #write-host "found match for $lastNameCheck"
                    $newUser.'username' = $newUser.'username' -replace "$lastNameCheck",$lastNameCheck.Substring(0,$k-1)
                    #write-host "It is now $($newUser.'username')"
                    break
                }
            }
        }
        
        #Student Template 7 & 8 do not use the full name and will always fit inside the 20 characters. This means we can use the generated username for the userprincipal as it will ALWAYS match.
        if (@(7,8) -contains $stuTemplate) {
            $newUser.'principalName' = $($newUser.'username') + $principalDomain
        } else {
            $newUser.'principalName' = $principalName -replace "$lastName","$($lastName)$([string]$i)"
        }

        $newUser.'homeDir' = "$($homeDirGradYR)\$($newUser.'username')"
        #now test and return
        #if ($(Get-AdUser -Filter "(SamAccountName -eq ""$($newUser.'username')"")") -or $(Get-AdUser -Filter "(UserPrincipalName -eq ""$($newUser.'principalName')"")")) {
        if ($(Get-AdUser -Filter "(SamAccountName -eq ""$($newUser.'username')"") -or (UserPrincipalName -eq ""$($newUser.'principalName')"")")) {
            Write-Host "$($newUser.'username'),$($newUser.'PrincipalName') is also not available."
            $errors++
            if ($errors -gt 10) {
                write-host "Error: You should never see this message unless you have a serious issue with username conflicts. You need to reconsider your student usernames." -ForeGroundColor Red
                [Environment]::Exit(1) #this is the only way I've found to kill the whole script once we are in the foreach-object loop
            }
        } else {
            return $newUser
        }
    }

}

function Add-ToADGroup ([String]$group,[Array]$users) {

    if ($($($users.GetType()).BaseType.Name) -ne 'Array') {
        Write-Host "Add-ToADGroup expects an array of users. Received $($($users.GetType()).BaseType.Name)"
        return $False
    }

    $members = @()
    #Fix for Groups Larger than Get-ADGroupMember can support.
    $groupDistinguishedName = Get-ADGroup -Identity $group | Select-Object -ExpandProperty DistinguishedName
    Get-ADUser -LDAPFilter "(&(objectCategory=user)(memberof=$($groupDistinguishedName)))" | Select-Object -Property SamAccountName | ForEach-Object { $members += $PSItem.'SamAccountName' }
    #Get-ADGroupMember -Identity $group | Select-Object SamAccountName | ForEach-Object { $members += $PSItem.'SamAccountName' }

    foreach ($i in $users) {
        #write-host "Checking if $i is in $group"
        #write-host $members; exit
        if (!($members -contains "$i")) {
            write-host "     - Adding $i to $group"
            Add-ADGroupMember -Identity $group -Members $i
        }
    }
    
    $members = $null    
    
}
function Set-ADGroupMembershipOnly ([String]$group,[Array]$users=@(),[Array]$excluded=@()) {
    
    Write-Host "Info: Verifying group memberships in $group"

    #Verify incoming $users variable as an array.
    if ($users.Length -eq 0 -or $users.Count -eq 0 ) { return }
    # if ($($($users.GetType()).BaseType.Name) -ne 'Array') {
    #     Write-Host "Error: Add-ToADGroup expects an array of users. Received $($($users.GetType()).BaseType.Name)" -ForegroundColor Red
    #     return $False
    # }
    
    $members = @()
    #Fix for Groups Larger than Get-ADGroupMember can support.
    $groupDistinguishedName = Get-ADGroup -Identity $group | Select-Object -ExpandProperty DistinguishedName
    Get-ADUser -LDAPFilter "(&(objectCategory=user)(memberof=$($groupDistinguishedName)))" | Select-Object -Property SamAccountName | ForEach-Object { $members += $PSItem.'SamAccountName' }
    #Get-ADGroupMember -Identity $group | Select-Object -Property SamAccountName | Where-Object { $members += $PSItem.'SamAccountName' }

    $removemembers = $members | Where-Object { $users -notcontains $PSItem }
    $removemembers = $removemembers | Where-Object { $excluded -notcontains $PSItem }
    if ($($removemembers | Measure-Object).count -ge 1) {
        write-host "Info:     - Removing $($removemembers -join ',') from $group"
        Remove-ADGroupMember -Identity $group -Members $removemembers -Confirm:$false
        #$numchanges++
    }

    $addmembers = $users | Where-Object { $members -notcontains $PSItem }
    $addmembers = $addmembers | Where-Object { $excluded -notcontains $PSItem }
    if ($($addmembers | Measure-Object).count -ge 1) {
        write-host "Info:     - Adding $($addmembers -join ',') to $group"
        Add-ADGroupMember -Identity $group -Members  $addmembers -Confirm:$false
        #$numchanges++
    }

    $members = $null; $removemembers = $null; $addmembers = $null

}

#Set the Ownership in the WBEMPath field for Groups. #Pretty sure this is wrong and not used!
function Set-ADGroupOwnershipOnly ([String]$group,[string]$buildingShortName,[string]$grade) {
    
    Write-Host "Info: Verifying group ownership for $group"

    #Verify incoming $users variable as an array.
    if ($users.Length -eq 0 -or $users.Count -eq 0 ) { return }
    # if ($($($users.GetType()).BaseType.Name) -ne 'Array') {
    #     Write-Host "Error: Add-ToADGroup expects an array of users. Received $($($users.GetType()).BaseType.Name)" -ForegroundColor Red
    #     return $False
    # }
    
    $members = @()
    $groupDistinguishedName = Get-ADGroup -Identity $group | Select-Object -ExpandProperty DistinguishedName
    Get-ADUser -LDAPFilter "(&(objectCategory=user)(memberof=$($groupDistinguishedName)))" | Select-Object -Property SamAccountName | ForEach-Object { $members += $PSItem.'SamAccountName' }
    #Get-ADGroupMember -Identity $group | Select-Object -Property SamAccountName | Where-Object { $members += $PSItem.'SamAccountName' }

    $removemembers = $members | Where-Object { $users -notcontains $PSItem }
    if ($($removemembers | Measure-Object).count -ge 1) {
        write-host "Info:     - Removing $($removemembers -join ',') from $group"
        Remove-ADGroupMember -Identity $group -Members $removemembers -Confirm:$false
        #$numchanges++
    }

    $addmembers = $users | Where-Object { $members -notcontains $PSItem }
    if ($($addmembers | Measure-Object).count -ge 1) {
        write-host "Info:     - Adding $($addmembers -join ',') to $group"
        Add-ADGroupMember -Identity $group -Members  $addmembers -Confirm:$false
        #$numchanges++
    }

    $members = $null; $removemembers = $null; $addmembers = $null

}

function Send-EmailNotification ([array]$mailto = @(), [string]$subject, [string]$body) {
    #something happens to the variables here that keeps Send-MailMessage from converting them. Moved to the actual script.
    write-host $mailto, $subject, $body

    . .\settings.ps1

    if ($mailto.Count -eq 0) {
        $mailto = @("$sendMailToEmail")
    }

    #write-host $mailto
    #write-host $sendMailToEmail

    # variables from settings.ps1
    # $sendMailNotifications = $True
    # $smtpAuth = $True
    # $smtpPasswordFile="C:\Scripts\emailpw.txt"
    # $sendMailToEmail = 'technology@gentrypioneers.com'
    # $sendMailFrom = 'technology@gentrypioneers.com'
    # $sendMailHost = 'smtp.gmail.com'
    # $sendMailPort = 587

    if (-Not($sendMailNotifications)) { Write-Host "Warning: Not configured to send email notifications."; return }
    
    #Send email via SMTP.
    if ($smtpAuth) {

        if (Test-Path ($smtpPasswordFile)) {
            $smtpPassword = Get-Content $smtpPasswordFile | ConvertTo-SecureString
        } else {
            Write-Host("SMTP Password file does not exist! [$smtpPasswordFile]. Please enter the password to be saved on this computer for emails.") -ForeGroundColor Yellow
            Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $smtpPasswordFile
            $smtpPassword = Get-Content $smtpPasswordFile | ConvertTo-SecureString
        }
        $mailCredentials = New-Object -Type System.Management.Automation.PSCredential -ArgumentList $sendMailFrom, $smtpPassword

        Send-MailMessage -From $sendMailFrom -to $mailto -Subject $subject `
            -Body $body -SmtpServer $sendMailHost -port $sendMailPort -UseSsl `
            -Credential $mailCredentials
    } else {
        Send-MailMessage -From $sendMailFrom -to $mailto -Subject $subject `
            -Body $body -SmtpServer $sendMailHost -port $sendMailPort -UseSsl
    }

}

function Get-NewPassword($student) {

    if ($UseDinoPassSimple -OR $UseDinoPassStrong) {
        try {
            if ($UseDinoPassSimple) {
                $dinoURL = "https://www.dinopass.com/password/simple"
            } elseif ($UseDinoPassStrong) {
                $dinoURL = "https://www.dinopass.com/password/strong"
            }
            $password = Invoke-RestMethod -Uri $dinoURL
            return $password
        } catch {
            #do nothing and continue with the randomly generated below.
        }
    }

    $words = @('think','cabin','trust','funny','prize','model','study','great','shine','world','light','unity','clear','first','piano','power','salad','phone','truth','depth','queen','chest','tooth','basis','world','guest','apple','entry','hotel','bread','night','steak','owner','pizza','skill','ratio','media','month','bonus','honey','uncle','movie','river','shirt','cheek','paper','photo','actor','video','youth','error','thing','buyer','topic','event','scene','heart')
    $specials = @('.','!','-','$','@')
    $generatedpassword = "$(Get-Random -InputObject $words)$(Get-Random -InputObject $specials)" + "$($student.Student_id)".substring("$($student.Student_id)".length - 4, 4)
    return $generatedpassword
}

function Remove-StringLatinCharacter {
    <#
.SYNOPSIS
    Function to remove diacritics from a string
.DESCRIPTION
    Function to remove diacritics from a string
.PARAMETER String
    Specifies the String that will be processed
.EXAMPLE
    Remove-StringLatinCharacter -String "L'été de Raphaël"
    L'ete de Raphael
.EXAMPLE
    Foreach ($file in (Get-ChildItem c:\test\*.txt))
    {
        # Get the content of the current file and remove the diacritics
        $NewContent = Get-content $file | Remove-StringLatinCharacter
        # Overwrite the current file with the new content
        $NewContent | Set-Content $file
    }
    Remove diacritics from multiple files
.NOTES
    Francois-Xavier Cat
    lazywinadmin.com
    @lazywinadmin
    github.com/lazywinadmin
    BLOG ARTICLE
        https://lazywinadmin.com/2015/05/powershell-remove-diacritics-accents.html
    VERSION HISTORY
        1.0.0.0 | Francois-Xavier Cat
            Initial version Based on Marcin Krzanowic code
        1.0.0.1 | Francois-Xavier Cat
            Added support for ValueFromPipeline
        1.0.0.2 | Francois-Xavier Cat
            Add Support for multiple String
            Add Error Handling
    .LINK
        https://github.com/lazywinadmin/PowerShell
#>
    [CmdletBinding()]
    PARAM (
        [Parameter(ValueFromPipeline = $true)]
        [System.String[]]$String
    )
    PROCESS {
        FOREACH ($StringValue in $String) {
            Write-Verbose -Message "$StringValue"

            TRY {
                [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($StringValue))
            }
            CATCH {
                $PSCmdlet.ThrowTerminatingError($PSItem)
            }
        }
    }
}

function Get-FNVHash {

    param(
        [string]$InputString
    )

    # Initial prime and offset chosen for 32-bit output
    # See https://en.wikipedia.org/wiki/Fowler–Noll–Vo_hash_function
    [uint32]$FNVPrime = 16777619
    [uint32]$offset = 2166136261

    # Convert string to byte array, may want to change based on input collation
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)

    # Copy offset as initial hash value
    [uint32]$hash = $offset

    foreach($octet in $bytes)
    {
        # Apply XOR, multiply by prime and mod with max output size
        $hash = $hash -bxor $octet
        $hash = $hash * $FNVPrime % [System.Math]::Pow(2,31)
    }
    return $hash
}

#Pull in custom/overriding functions.
if (Test-Path $currentPath\z_functionsCustom.ps1) {
    . $currentPath\z_functionsCustom.ps1
}

function Connect-Database {
    param (
        [Parameter(Mandatory=$true)]$database #this could be a hashtable or a string.
    )

    #Default to SQLite
    if ($database.GetType().Name -eq 'String') {
        try {
            Open-SQLiteConnection -DataSource $database
            return $True
        } catch {
            return $False
        }
    }

    $dbcredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $database['username'],(ConvertTo-SecureString -AsPlainText "$($database['password'])" -Force)

    #MySQL or MSSQL
    if ($database.dbtype -eq 'mysql') {
        try {
            Open-MySqlConnection -Server $database.hostname -Database $database.dbname -Credential $dbcredentials
            return $True
        } catch {
            return $False
        }
    } elseif ($database.dbtype -eq 'mssql') {
        try {
            Open-SqlConnection -Server $database.hostname -Database $database.dbname -UserName $database.username -Password $database.password
            return $True
        } catch {
            return $False
        }
    }

    #we should have returned something by now. Otherwise just return False.
    return $False

}

Function Out-DataTable {
  $dt = new-object Data.datatable  
  $First = $true  
 
  foreach ($item in $input){  
    $DR = $DT.NewRow()  
    $Item.PsObject.get_properties() | foreach {  
      if ($first) {  
        $Col =  new-object Data.DataColumn  
        $Col.ColumnName = $_.Name.ToString()  
        $DT.Columns.Add($Col)       }  
      if ($_.value -eq $null) {  
        $DR.Item($_.Name) = "[empty]"  
      }  
      elseif ($_.IsArray) {  
        $DR.Item($_.Name) =[string]::Join($_.value ,";")  
      }  
      else {  
        $DR.Item($_.Name) = $_.value  
      }  
    }  
    $DT.Rows.Add($DR)  
    $First = $false  
  } 
 
  return @(,($dt))
 
}