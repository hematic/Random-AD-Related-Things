If(-not(Test-DotNetFrameWork35)) { "Requires .NET Framework 3.5" ; exit }

#region Script Variables

# Set this value to the number of users with large tokens that you want to report on.
$TopUsers = 200

# Set this to the size in bytes that you want to capture the user information for the report.
#Don't set this too low. Crashed my desktop when it ran out of memory :(
$TokensSizeThreshold = 6000

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

# Set this value to true if you want to output to the console
$ConsoleOutput = $True

# Set this value to true if you want a summary output to the console when the script has completed.
$OutputSummary = $True

# Set this value to true to use the tokenGroups attribute
$UseTokenGroups = $True

# Set this value to true to use the GetAuthorizationGroups() method
$UseGetAuthorizationGroups = $False

# Set the script path
$ScriptPath = 'C:\temp'
$ReferenceFile = $ScriptPath + "\KerberosTokenSizeReport.csv"

$array = @()
$TotalUsersProcessed = 0
$UserCount = 0
$GroupCount = 0
$LargestTokenSize = 0
$TotalGoodTokens = 0
$TotalTokensBetween8and12K = 0
$TotalLargeTokens = 0
$TotalVeryLargeTokens = 0

#endregion
#region Function Declarations

Function Get-UserPrincipal($cName, $cContainer, $userName){
  $dsam = "System.DirectoryServices.AccountManagement" 
  $rtn = [reflection.assembly]::LoadWithPartialName($dsam)
  $cType = "domain" #context type
  $iType = "SamAccountName"
  $dsamUserPrincipal = "$dsam.userPrincipal" -as [type]
  $principalContext = new-object "$dsam.PrincipalContext"($cType,$cName,$cContainer)
  $dsamUserPrincipal::FindByIdentity($principalContext,$iType,$userName)
}

Function Test-DotNetFrameWork35{
 Test-path -path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5'
}

#endregion
#region get all Non-System Accounts

$ADRoot = ([System.DirectoryServices.DirectoryEntry]"LDAP://RootDSE")
$DefaultNamingContext = $ADRoot.defaultNamingContext

# Derive FQDN Domain Name
$TempDefaultNamingContext = $DefaultNamingContext.ToString().ToUpper()
$DomainName = $TempDefaultNamingContext.Replace(",DC=",".")
$DomainName = $DomainName.Replace("DC=","")

# Create an LDAP search for all enabled users not marked as criticalsystemobjects to avoid system accounts
$ADFilter = "(&(objectClass=user)(objectcategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(!(isCriticalSystemObject=TRUE))(!name=IUSR*)(!name=IWAM*)(!name=ASPNET))"
# There is a known bug in PowerShell requiring the DirectorySearcher properties to be in lower case for reliability.
$ADPropertyList = @("distinguishedname","samaccountname","useraccountcontrol","objectsid","sidhistory","primarygroupid","lastlogontimestamp","memberof")
$ADScope = "SUBTREE"
$ADPageSize = 1000
$ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DefaultNamingContext)") 
$ADSearcher = New-Object System.DirectoryServices.DirectorySearcher 
$ADSearcher.SearchRoot = $ADSearchRoot
$ADSearcher.PageSize = $ADPageSize 
$ADSearcher.Filter = $ADFilter 
$ADSearcher.SearchScope = $ADScope
if ($ADPropertyList) {
  foreach ($ADProperty in $ADPropertyList) {
    [Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
  }
}
$Users = $ADSearcher.Findall()
$UserCount = $users.Count

#endregion

If ($UserCount -ne 0) {
    foreach($user in $users) {
        
        $lastLogonTimeStamp = ""
        $lastLogon = ""
        $UserDN = $user.Properties.distinguishedname[0]
        $samAccountName = $user.Properties.samaccountname[0]
        
        If (($user.Properties.lastlogontimestamp | Measure-Object).Count -gt 0) {
            $lastLogonTimeStamp = $user.Properties.lastlogontimestamp[0]
            $lastLogon = [System.DateTime]::FromFileTime($lastLogonTimeStamp)
            if ($lastLogon -match "1/01/1601") {
                $lastLogon = "Never logged on before"
            }
        } 
        Else{
            $lastLogon = "Never logged on before"
        }
        
        $OU = $user.GetDirectoryEntry().Parent
        $OU = $OU -replace ("LDAP:\/\/","")

        # Get user SID
        $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($user.Properties.objectsid[0], 0)
        $userSID = $arruserSID.Value

        # Get the SID of the Domain the account is in
        $AccountDomainSid = $arruserSID.AccountDomainSid.Value

        # Get User Account Control & Primary Group by binding to the user account
        $objUser = [ADSI]("LDAP://" + $UserDN)
        $UACValue = $objUser.useraccountcontrol[0]
        $primarygroupID = $objUser.PrimaryGroupID
        
        # Primary group can be calculated by merging the account domain SID and primary group ID
        $primarygroupSID = $AccountDomainSid + "-" + $primarygroupID.ToString()
        $primarygroup = [adsi]("LDAP://<SID=$primarygroupSID>")
        $primarygroupname = $primarygroup.name
        $objUser = $null

        # Get SID history
        $SIDCounter = 0
        if ($user.Properties.sidhistory -ne $null) {
            foreach ($sidhistory in $user.Properties.sidhistory) {
                $SIDHistObj = New-Object System.Security.Principal.SecurityIdentifier($sidhistory, 0)
                $SIDCounter++
            }
        }
        
        $SIDHistObj = $null
        $TotalUsersProcessed ++
    
        If ($ProgressBar) {
            Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $samAccountName) -PercentComplete (($TotalUsersProcessed/$UserCount)*100)
        }

        # Use TokenGroups Attribute
        If ($UseTokenGroups) {
            
            $UserAccount = [ADSI]"$($User.Path)"
            $UserAccount.GetInfoEx(@("tokenGroups"),0) | Out-Null
            $ErrorActionPreference = "continue"
            $error.Clear()
            $groups = $UserAccount.GetEx("tokengroups")
            
            if ($Error) {
                Write-Warning "  Tokengroups not readable"
                $Groups=@()
            }
        
            $GroupCount = 0

            # Note that the tokengroups includes all principals, which includes siDHistory, so we need
            # to subtract the sIDHistory count to correctly report on the number of groups in the token.
            $GroupCount = $groups.count - $SIDCounter

            $SecurityDomainLocalScope = 0
            $SecurityGlobalInternalScope = 0
            $SecurityGlobalExternalScope = 0
            $SecurityUniversalInternalScope = 0
            $SecurityUniversalExternalScope = 0

            foreach($token in $groups) {
            
                $principal = New-Object System.Security.Principal.SecurityIdentifier($token,0)
                $GroupSid = $principal.value
                $grp = [ADSI]"LDAP://<SID=$GroupSid>"
        
                if ($grp.Path -ne $null) {
                    $grpdn = $grp.distinguishedName.tostring().ToLower()
                    $grouptype = $grp.groupType.psbase.value

                    switch -exact ($GroupType) {
                        "-2147483646"   { 
                            # Global security scope 
                            if ($GroupSid -match $DomainSID){
                                $SecurityGlobalInternalScope++
                            } 
                            else { 
                                # Global groups from others.
                                $SecurityGlobalExternalScope++
                            } 
                        } 
                        "-2147483644"   { 
                            # Domain Local scope 
                            $SecurityDomainLocalScope++
                        } 
                        "-2147483643"   { 
                            # Domain Local BuildIn scope
                            $SecurityDomainLocalScope++
                        }
                        "-2147483640"   { 
                            # Universal security scope 
                            if ($GroupSid -match $AccountDomainSid){ 
                                $SecurityUniversalInternalScope++ 
                            } 
                            else{ 
                                # Universal groups from others.
                                $SecurityUniversalExternalScope++ 
                            } 
                        } 
                    }
                }
            } 
        }

        # Use GetAuthorizationGroups() Method
        If ($UseGetAuthorizationGroups) {

            $userPrincipal = Get-UserPrincipal -userName $SamAccountName -cName $DomainName -cContainer "$OU"

            $GroupCount = 0
            $SecurityDomainLocalScope = 0
            $SecurityGlobalInternalScope = 0
            $SecurityGlobalExternalScope = 0
            $SecurityUniversalInternalScope = 0
            $SecurityUniversalExternalScope = 0

            # Use GetAuthorizationGroups() for Indirect Group MemberShip, which includes all Nested groups and the Primary group
            Try {
                $groups = $userPrincipal.GetAuthorizationGroups() | select SamAccountName, GroupScope, SID
                $GroupCount = $groups.count

                foreach ($group in $groups) {
                
                    $GroupSid = $group.SID.value

                    switch ($group.GroupScope)
                                                                                                                                                                                                                                                                                                            {
                    "Local" {
                    # Domain Local & Domain Local BuildIn scope
                    $SecurityDomainLocalScope++
                    }
                    "Global" {
                    # Global security scope 
                    if ($GroupSid -match $DomainSID) {
                        $SecurityGlobalInternalScope++
                    } else { 
                        # Global groups from others.
                        $SecurityGlobalExternalScope++
                    }
                    }
                    "Universal" {
                    # Universal security scope 
                    if ($GroupSid -match $AccountDomainSid) {
                        $SecurityUniversalInternalScope++
                    } else {
                        # Universal groups from others.
                        $SecurityUniversalExternalScope++
                    }
              }
                }
                }
            }
                    Catch {
            write-host "Error with the GetAuthorizationGroups() method: $($_.Exception.Message)" -ForegroundColor Red
        }
        }

        If ($ConsoleOutput) {
            Write-Host -ForegroundColor green "Checking the token of user $SamAccountName in domain $DomainName"
            Write-Host -ForegroundColor green "There are $GroupCount groups in the token."
            Write-Host -ForegroundColor green "- $SecurityDomainLocalScope are domain local security groups."
            Write-Host -ForegroundColor green "- $SecurityGlobalInternalScope are domain global scope security groups inside the users domain."
            Write-Host -ForegroundColor green "- $SecurityGlobalExternalScope are domain global scope security groups outside the users domain."
            Write-Host -ForegroundColor green "- $SecurityUniversalInternalScope are universal security groups inside the users domain."
            Write-Host -ForegroundColor green "- $SecurityUniversalExternalScope are universal security groups outside the users domain."
            Write-host -ForegroundColor green "The primary group is $primarygroupname."
            Write-host -ForegroundColor green "There are $SIDCounter SIDs in the users SIDHistory."
            Write-Host -ForegroundColor green "The current userAccountControl value is $UACValue."
        }

        $TrustedforDelegation = $false
        if ((($UACValue -bor 0x80000) -eq $UACValue) -OR (($UACValue -bor 0x1000000) -eq $UACValue)) {
            $TrustedforDelegation = $true
        }

        # Calculate the current token size, taking into account whether or not the account is trusted for delegation or not.
        $TokenSize = 1200 + (40 * ($SecurityDomainLocalScope + $SecurityGlobalExternalScope + $SecurityUniversalExternalScope + $SIDCounter)) + (8 * ($SecurityGlobalInternalScope  + $SecurityUniversalInternalScope))
        if ($TrustedforDelegation -eq $false) {
            If ($ConsoleOutput) {
                Write-Host -ForegroundColor green "Token size is $Tokensize and the user is not trusted for delegation."
            }
        } 
        else {
            $TokenSize = 2 * $TokenSize
            If ($ConsoleOutput) {
                Write-Host -ForegroundColor green "Token size is $Tokensize and the user is trusted for delegation."
            }
        }

        If ($TokenSize -le 12000) {
            $TotalGoodTokens ++
            If ($TokenSize -gt 8192) {
                $TotalTokensBetween8and12K ++
            }
        } 
        elseIf ($TokenSize -le 48000) {
            $TotalLargeTokens ++
        } 
        else {
            $TotalVeryLargeTokens ++
        }

        If ($TokenSize -gt $LargestTokenSize) {
            $LargestTokenSize = $TokenSize
            $LargestTokenUser = $SamAccountName
        }

        If ($TokenSize -ge $TokensSizeThreshold) {
        
                                                                    $obj = New-Object -TypeName PSObject -Property @{
        "Domain"               = $DomainName
        "SamAccountName"       = $SamAccountName
        "TokenSize"            = $TokenSize
        "Memberships"          = $GroupCount
        "DomainLocal"          = $SecurityDomainLocalScope
        "GlobalInternal"       = $SecurityGlobalInternalScope
        "GlobalExternal"       = $SecurityGlobalExternalScope
        "UniversalInternal"    = $SecurityUniversalInternalScope
        "UniversalExternal"    = $SecurityUniversalExternalScope
        "SIDHistory"           = $SIDCounter
        "UACValue"             = $UACValue
        "TrustedforDelegation" = $TrustedforDelegation
        "LastLogon"            = $lastLogon
        }
            $array += $obj
        }

        If ($ConsoleOutput) {
            $percent = "{0:P}" -f ($TotalUsersProcessed/$UserCount)
            write-host -ForegroundColor green "Processed $TotalUsersProcessed of $UserCount user accounts = $percent complete."
            Write-host " "
        }

        If ($OutputSummary) {
            Write-Host -ForegroundColor green "Summary:"
            Write-Host -ForegroundColor green "- Processed $UserCount user accounts."
            Write-Host -ForegroundColor green "- $TotalGoodTokens have a calculated token size of less than or equal to 12000 bytes."
            
            If ($TotalGoodTokens -gt 0) {
            Write-Host -ForegroundColor green "  - These users are good."
            }
            If ($TotalTokensBetween8and12K -gt 0) {
                Write-Host -ForegroundColor green "  - Although $TotalTokensBetween8and12K of these user accounts have tokens above 8K and should therefore be reviewed."
            }
            Write-Host -ForegroundColor green "- $TotalLargeTokens have a calculated token size larger than 12000 bytes."
            If ($TotalLargeTokens -gt 0) {
                Write-Host -ForegroundColor green "  - These users will be okay if you have increased the MaxTokenSize to 48000 bytes.`n  - Consider reducing direct and transitive (nested) group memberships."
            }
            Write-Host -ForegroundColor red "- $TotalVeryLargeTokens have a calculated token size larger than 48000 bytes."
            If ($TotalVeryLargeTokens -gt 0) {
                Write-Host -ForegroundColor red "  - These users will have problems. Do NOT increase the MaxTokenSize beyond 48000 bytes.`n  - Reduce the direct and transitive (nested) group memberships."
            }
            Write-Host -ForegroundColor green "- $LargestTokenUser has the largest calculated token size of $LargestTokenSize bytes in the $DomainName domain."
        }

        # Write-Output $array | Format-Table
        $array | Sort-Object TokenSize -descending | select-object -first $TopUsers | export-csv -notype -path "$ReferenceFile" -Delimiter ';'

        # Remove the quotes
        (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
}
}