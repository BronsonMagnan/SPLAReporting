#Note on service accounts:
#The scheduled task this report runs under, needs a service account which has the following rights
#Hyper-V Administrator and Remote Management on each hypervisor
#Local Admin on each SQL Server
#The ability to read active directory group memberships.


#This flag controls logging to console, the destination email address, and the date this runs on so you can debug the reporting
#In this example, contoso is the tenant environment and fabrikam is the service provider
$debugmode = $false
$debugEmailAddress = "programmer@fabrikam.com"
$reportingEmailAddress = "splareporting@fabrikam.com"
$SMTPServer = "exch201901.corp.contoso.com"
$fromEmailAddress = "automation@contoso.com"
$tenantName = "Contoso"


#This OU must contain all the groups group that represents each product SKU, which the billiable users chain membership into.
#The script does not detect products, but depends on you using a group membership structure that ties to users to the SKUs.
$SPLAGroupOUContainer = "OU=SPLA SAL Reporting Groups,OU=Groups,OU=Datacenter,DC=Corp,DC=Contoso,DC=Com"
<#
    Examples of groups to put in this OU
    SPLA-OfficeProPlus
    SPLA-ProjectStandard
    SPLA-RemoteDesktopServicesSAL
    SPLA-UserExperienceVirtualization

    Note the Description of the Group in AD, is what will appear for the SKU in the report.
#>

#This AD Group contains all the SQL Server Standard Edition Virtual Machine accounts
#Put this in a different OU than the SAL groups, examples "OU=SPLA Core Reporting Groups,..."
$SqlStandard = "SPLA-SQLStandardServers"

#This varible is a regex to exclude administrator accounts
#The location we look for users to count, is in the default users container.
#Assumes you have moved the Default User location to a new OU, and not the default user container.
$adminfilter = "administrator|testuser1|testuser2|tenantAdmin1"

#Run the report on the first day of the month, and run again on the 25th of the month so you can catch any stale accounts before the next reporting deadline.
$dayofMonth = 1
$dayofMonth2 = 25

#---End of Environment specific definitions---

#only run on the specified day of the month, the task will run every night. always run in DebugMode
$today = [datetime]::Today
if (($today.day -eq $dayofMonth) -or ($today.day -eq $dayofMonth2) -or  ($debugmode))
{
#The whole script is wrapped in this IF block

$GlobalLog = @()
function write-log () {
    param($message)
    $Global:GlobalLog += "$(get-date -DisplayHint Time): $message"
    Write-Host "$(get-date -DisplayHint Time): $message"
}

#Calculate the billing period
$date = get-date
$startofmonth = Get-Date $date.AddMonths(-1) -day 1 -hour 0 -minute 0 -second 0
$endofmonth = (($startofmonth).AddMonths(1).AddSeconds(-1))


class hypervisor {
    [string]$computername
    [int]$vmcount
    [int]$processorCount
    [int]$coreCount
    [string]$serverEdition
    [boolean]$clusterNode
    [string]$SKUType
    [int]$SKUQty
    [int]$minCores = 8
    [int]$minProcs = 1
    [int]$skusize = 2
    [boolean]$notahypervisor 

    hypervisor($name) {
        $this.computername = $name
        Write-Log "Querying $($this.computername) for VM count"
        try {
            $this.vmcount = Invoke-Command -ComputerName $this.computername -ScriptBlock { return (get-vm | measure).count} -ErrorAction Stop
        } catch {
            $hypervRole = (get-windowsfeature -ComputerName $this.computername -name Hyper-V).installed
            $hypervPS = (get-windowsfeature -ComputerName $this.computername -name Hyper-V-Powershell).installed
            if (($hypervRole) -and (-not $hypervPS)) {
            }
            if (-not $hypervRole) {
                $this.vmcount = 0
                $this.notahypervisor = $true
            } elseif (-not $hypervPS) {
                write-error "Hyper-V-Powershell module missing on hypervisor $($this.computername)"
                exit
            } else {
                Write-Error "FATAL ERROR $($this.computer)"
                exit
            }
        }
        Write-Log "Querying $($this.computername) for Processor Data"
        $processors = Get-CimInstance -ComputerName $this.computername -ClassName Win32_Processor
        if ($processors -ne $null) {
            $this.processorCount = $($processors|Measure-Object).count
            $this.coreCount = $($processors.numberofCores | Select-Object -First 1)
        }
        Write-Log "Querying $($this.computername) for Operating System information"
        $serverVersion = (Get-CimInstance -ComputerName $this.computername -ClassName Win32_OperatingSystem).caption
        if ($serverVersion -ne $null) {
            if ($serverVersion.toupper() -match "DATACENTER") {
                $this.serverEdition = "Datacenter"
            } elseif ($serverVersion.toupper() -match "STANDARD") {
                $this.serverEdition = "Standard"
            } else {
                $this.serverEdition = "UNKNOWN"
            }
        }
        Write-Log "Querying $($this.computername) for Cluster Service"
        $SystemServices = Get-CimInstance -computername $this.computername -ClassName Win32_SystemServices
        if ($SystemServices -ne $null) {
            $clusterService = $null
            $clusterService = ($SystemServices | where {$_.partcomponent -match "CLUSSVC"})
            if ($clusterService -ne $null) {
                $this.clusterNode = $true
            } else {
                $this.clusterNode = $false
            }
        }
        Write-Log "Calculating SPLA for $($this.computername)"
        $this.calcSpla()
    } # end constructor

    [string]plural ([int]$what) {
        if ($what -eq 1) {
            return "" 
        } else { 
            return "s"
        }
    }
    [string]clusterStatus () {
        if ($this.clusterNode) {
            return "Clustered"
        } else {
            return "Standalone"
        }
    }
    [string]hostingStatement() {
        if ($this.vmcount -eq 0) {
            if ($this.notahypervisor) {
                return "not a hypervisor"
            } else {
                return "not hosting virtual machines"
            }
        } else {
            return "hosting $($this.vmcount) virtual machine$($this.plural($this.vmcount))"
        }
    }
    [string]reportStats () {
        return "$($this.computername), is $($this.clusterStatus()), and has $($this.processorCount) processor$($this.plural($this.processorCount)) with $($this.coreCount) cores, running Windows Server $($this.serverEdition) edition, and is $($this.hostingStatement())"
    }
    [string]reportSpla() {
        return "$($this.computername) will be assigned $($this.SKUQty) license$($this.plural($this.SKUQty)) of SKU $($this.SKUType)"
    }
    [void]subStd() {
        #report DC
        $this.SKUType = "Standard"
        #the Std sku is processors min 1, * cores, min 8, /2 because they are 2 core packs all that times the number of virtual machines
        $this.SKUQty = ($this.vmcount) * ([math]::max($this.minProcs,$this.processorCount)) * ([math]::max($this.minCores,$this.coreCount)) / $this.skuSize
    }
    [void]subDC() {
        #report DC
        $this.SKUType = "Datacenter"
        #the DC sku is processors min 1, * cores, min 8, /2 because they are 2 core packs
        $this.SKUQty = ([math]::max($this.minProcs,$this.processorCount)) * ([math]::max($this.minCores,$this.coreCount)) / $this.skuSize
    }
    [void]calcSpla() {
        #if the server is not a hypervisor
        if ($this.notahypervisor) {
            #use the DC calculation, then replace the SKU with the server edition
            $this.subDC()
            $this.SKUType = $this.serverEdition
        } else {
            #if the server is running 0 vms, no payment is due
            if (($this.vmcount -eq 0) -and ($this.clusterNode -eq $false)) {
                $this.SKUQty = 0
                $this.SKUType = $this.serverEdition
            } else {
                if ($this.clusterNode -eq $true) {
                    #clusterNode - report DC, dont be stupid, it gets astronomical to report std here.
                    $this.subDC()
                } else {
                    #standalone
                    if ($this.serverEdition.toupper() -eq "DATACENTER") {
                        $this.subDC()
                    } elseif ($this.serverEdition.toupper() -eq "STANDARD") {
                        if ($this.vmcount -le 4) {
                            #Report Std because it is lightly virtualized
                            $this.subStd()
                        } else {
                            #Report DC because more economical
                            $this.subDC()
                        }
                    } else {
                        $this.SKUType = "Unknown"
                        $this.SKUQty = -99999
                    } 
                } #if clustertest
            } #if more vms than 0
        } #not a hypervisor test
    }


}





#SQL Standard by 2CoreCountPack
$sqlservers = (get-adgroup -Identity $SqlStandard | Get-ADGroupMember)
$corecount = 0
foreach ($server in $sqlservers) 
{
    #Change to counting logical processors because of the change in SMT on virtual machines in the datacenter.
    $thisServer = 	(get-ciminstance -computername $server.name -classname Win32_Processor).NumberofLogicalProcessors
    write-log "SQL Server $($server.name) has $($thisServer) logical processors and will be licensed for $($thisServer/2) 2-core packs"
	$corecount += $thisServer
}
$SqlStdCore = ($Corecount/2)

#Hyper Licensing
$Dnames = Get-ADObject -Filter 'ObjectClass -eq "serviceConnectionPoint" -and Name -eq "Microsoft Hyper-V"' | Select-Object -ExpandProperty DistinguishedName
$ServerNames = ($dnames.split(',') | Where-Object {$_ -ne "CN=Microsoft Hyper-V"}| Where-Object {$_ -match "CN=*"}).replace("CN=","")
#Remove Nested Hypervisors
$nestedHyperV=@()
foreach ($server in $servernames) {
    Write-Log "Checking if $server is a nested hypervisor"
    $ComputerInfo = $null
    $ComputerInfo = Get-CimInstance -ComputerName $server -ClassName win32_computersystem
    $IsVM = $($ComputerInfo.Model -eq "Virtual Machine")
    if ($isVM) {
        Write-Log "$server is a nested hypervisor - dont count in licensing"
        $nestedHyperV += $server
    } else { 
        Write-Log "$server is a physical installation - $($ComputerInfo.Model)"
    }
}
$ServerNames = $ServerNames | where {$_ -notin $nestedHyperV}


$hyperv = @()
foreach ($server in $ServerNames) {
    Write-Log "Processing $server"
    $hyperv += [hypervisor]::new($server)
}
foreach ($system in $hyperv) {
    write-log $system.reportStats()
    write-log $system.reportSpla()
}
$StdSkuCount = ($hyperv | where {$_.SkuType.toupper() -eq "STANDARD"} | select-object -ExpandProperty SkuQty | measure-object -sum).Sum
$DCSkuCount = ($hyperv | where {$_.SkuType.toupper() -eq "DATACENTER"} | select-object -ExpandProperty SkuQty | measure-object -sum).Sum


#Get the group list in the SPLA OU


#We need to calculate any user accounts that were not active in the last billing perioud
$userlocation = (get-addomain).UsersContainer
$UsersToCheck = Get-ADUser -SearchBase $userlocation -Filter * -Properties "DisplayName", "UserPrincipalName", "msDS-UserPasswordExpiryTimeComputed", "LastLogonTimeStamp", "Enabled", "Surname", "GivenName", "whenCreated", "PasswordNeverExpires"
write-log "Checking User Accounts in $userlocation"
$UsersToFilter = @()
foreach ($UserAccount in $UsersToCheck) {
    $flagNoLogonLastMonth = $false
    $flagAccountDisabled = $false
    $flagVeryNewAccount = $false
    $flagPasswordExpiredBeforeBillingPeriod = $false

    write-log "User: $($UserAccount.DisplayName)"

    #Check if the account is enabled of disabled
    if (-not $UserAccount.Enabled) {
        write-log "-Status: Disabled"
        $flagAccountDisabled = $true
    } else { 
        write-log "-Status: Enabled"
    }

    #Check for creation after the billing period
    $UserAccountCreationDate = [datetime]$userAccount."whenCreated"
    if ($UserAccountCreationDate -gt $endofmonth) { 
        write-log "-Created: $UserAccountCreationDate - after the billing period ending on $endofmonth"
        $flagVeryNewAccount = $true
    } else {
        write-log "-Created: $UserAccountCreationDate"
    }

    #Check for the user account last logon before the start of the billing period
    $UserLastLogon = [datetime]::FromFileTime($UserAccount.LastLogonTimeStamp)
    if ($UserLastLogon -lt $startofmonth) {
        write-log "-LastLogon: $UserLastLogon - Account has not logged in during the billing period starting on $startofmonth"
        $flagNoLogonLastMonth = $true
    } else {
        write-log "-LastLogon: $UserLastLogon"
    }

    if (-not $UserAccount."PasswordNeverExpires") {
        $UserPasswordExpirationDate = [datetime]::FromFileTime($UserAccount."msDS-UserPasswordExpiryTimeComputed")
        if ($UserPasswordExpirationDate -lt $startofmonth) {
            $flagPasswordExpiredBeforeBillingPeriod = $true
            write-log "-PasswordExpiration: $UserPasswordExpirationDate - this password is expired"
        } else {
            write-log "-PasswordExpiration: $UserPasswordExpirationDate"
        }
    } else {
            write-log "-PasswordExpiration: $UserPasswordExpirationDate - Password Does NOT expire"
    }

    <# Logic to exclude users
    $flagNoLogonLastMonth
    $flagAccountDisabled
    $flagVeryNewAccount
    $flagPasswordExpiredBeforeBillingPeriod

    now the rules are that we must bill someone if the COULD have used the solution during the billing period.
    therefore
        an account created after the billing period is exempt.
        an account that is disabled is exempt.
        an account whose password expired before the billing period is exempt.
        no logon last month is an indicator that we should have taken action to block the account.
    #>
    if ($flagVeryNewAccount -or $flagAccountDisabled -or $flagPasswordExpiredBeforeBillingPeriod) {
        write-log "-BillableStatus: not billable"
        $UsersToFilter += $UserAccount
    } else {
        if ($flagNoLogonLastMonth) {
            write-log "-BillableStatus: billable (but could have been prevented with proper user management)"
        } else {
            write-log "-BillableStatus: billable"
        }
    }
}

write-log "Summary List of Users to Filter"
if ($UsersToFilter) {
    foreach ($UserToFilter in $UsersToFilter) { 
        write-log "User: $($UserToFilter."DisplayName")"
    }
} else {
    write-log "No filterable users based on usage rules"
}

#assemble the regex filter.
if ($UsersToFilter) {
    $userFilter = ""
    foreach ( $UserToFilter in $UsersToFilter ) { 
        $userFilter += $UserToFilter.SamAccountName
        $userFilter += "|"
    }
    if ($userFilter -ne "") { 
        if ($userFilter.Substring($userfilter.length -1 ,1) = "|") { 
            $userFilter = $userFilter.Substring(0,$userfilter.Length -1) 
        }
    }
}

$UserSkuList = @()
$SPLAGroupList = Get-ADGroup -filter * -Properties Description | where {$_.DistinguishedName -like "*$SPLAGroupOUContainer"} | sort name

foreach ($Group in $SPLAGroupList) {
    if ($UsersToFilter) {
        $users = Get-ADGroupMember $Group.name -Recursive | where {$_.SamAccountName -notmatch $adminFilter} | where {$_.samAccountName -notmatch $userFilter}
    } else {
        $users = Get-ADGroupMember $Group.name -Recursive | where {$_.SamAccountName -notmatch $adminFilter}
    }
    $UserSkuList += @{
        "SKU"=$Group.description;
        "Count"=$users.count;
        "List"=($users).name | sort 
    }
}

#Sort is very important
$UserSkuList = $UserSkuList | sort SKU

$templateHeader = @"
 <html>
 <head>
 <title>SPLA Inventory: $((Get-ADdomain).dnsroot)</title>
 </head>
 <body>
 <table>
 <tr><td span ="2"><u>Usage</u></td></tr>
"@

$templateCenter = ""
foreach ($UserSku in $UserSkuList) {
    $templateCenter += "<tr><td>$($UserSku.Sku):</td><td> $($UserSku.Count)</td></tr>"
}


$templateDetails = "<table border='1'>"
$templateDetails += "<tr><td>Users</td>"

foreach ($UserSku in $UserSkuList) {
    $templateDetails += "<td>$($UserSku.SKU)</td>"
}
$templateDetails += "</tr>"
#Get the user list:
$UserList = ($UserSkuList).list| sort -Unique 
foreach ($User in $UserList) {
$templateDetails += "<tr><td>$User</td>"
    foreach ($sku in $UserSkuList) {
        if ($sku.list -contains $user) {
            $templateDetails += "<td>X</td>"
        } else {
            $templateDetails += "<td></td>"
        }
    }
$templateDetails += "</tr>"
}
$templateDetails += "</table>"

$templateFooter = @"
 <tr><td>   SQL Std 2 Core Pack:</td><td> $SqlStdCore         </td></tr>
 <tr><td>   Server 2016 Std SKU:</td><td> $StdSkuCount        </td></tr>
 <tr><td>    Server 2016 DC SKU:</td><td> $DCSkuCount         </td></tr>
 </table>
"@ 


$template = ""
$template += $templateHeader
$template += $templateCenter
$template += $templateFooter
$template += $templateDetails

$template += "</body> </html><br>Logfile: <br>"

write-log "Username: $(&whoami)"
write-log "System: $ENV:COMPUTERNAME"
write-log "DebugMode: $debugmode"

foreach ($line in $GlobalLog) {
    $template += "$line<br>"
}


Write-Output $template


if ($debugmode) { 
    $mailaddress = $debugEmailAddress
} else {
    $mailaddress = $reportingEmailAddress
}

$MailParams = @{
    "To"            = $mailaddress;
    "Subject"       = "SPLA Detail: $tenantName";
    "Body"          = $template;
    "SmtpServer"    = $SMTPServer;
    "BodyAsHtml"    = $true;
    "From"          = $fromEmailAddress
    "Port"          = "25";
}

Send-MailMessage @MailParams

}
