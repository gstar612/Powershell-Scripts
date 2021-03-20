<#
Author: Harrison Muncaster
Date: 01/10/2021
Purpose: Import CSV of contacts needing to be created in Active Directory.
If contact already exist, has no mail address, or invalid mail domain then skip.
Otherwise, set company variable based off of mail domain, set name in {FullName - Company} format & create contact object in target OU.
#>


Import-Module ActiveDirectory


$ContactsToImport = Import-Csv -path $env:FILEPATH
$ExistingContacts = Get-ADObject -Filter 'objectClass -eq "contact"' -SearchBase $env:SEARCHBASE -Properties Mail
$TotalCount = $ContactsToImport.count
$Counter = 1

$ContactsToImport | % {
    
    write-host "Processing $Counter of $TotalCount"
    $Counter++
    
    if (!$_.mail) {
        write-host "No Email Present - Skipping"
        return
    }

    if ($_.targetAddress) {
        $targetAddress = ($_.targetAddress).Split(':')[-1]
    } else {
        $targetAddress = $_.mail
      }

    if ($targetAddress -in $ExistingContacts.Mail) {
        Write-Host "Found a match! Skipping, Already created"
        return

    } else {

        # Set $Company variable based off of email domain
        $Company = switch -Wildcard ($targetAddress) {
            {$_ -like "*@disney.com*"} {"Disney"; Break}
            {$_ -like "*@espn.com*"} {"ESPN"; Break}
            {$_ -like "*@natgeo.com*"} {"NatGeo"; Break}
            {$_ -like "*@abc.com*"} {"ABC"; Break}
            {$_ -like "*@truex.com*"} {"TRUEX"; Break}
            default {$null}
        }

        if (!$Company) {
            Write-Host "No Company Present - Skipping"
            return
        }

        $Name = switch ($_) {
            {$_.GivenName -and $_.sn} {$_.GivenName + " " + $_.sn; Break}
            {$_.GivenName -and !$_.sn} {$_.GivenName; Break}
            {!$_.GivenName -and $_.sn} {$_.mail; Break}
            {!$_.GivenName -and !$_.sn} {$targetAddress; Break}
        }
        $Name = $Name.trim()
        $Name += " - $Company"

        # Set $Params variable for contact that will be used in new-adobject call
        $Params = switch ($_) {
            {!$_.GivenName -and $_.sn} {@{sn=$_.sn;displayName=$name;company=$company;mail=$_.mail;targetAddress=$targetAddress}; Break}
            {!$_.sn -and $_.GivenName} {@{givenName=$_.GivenName;displayName=$name;company=$company;mail=$_.mail;targetAddress=$targetAddress}; Break}
            {!$_.GivenName -and !$_.sn} {@{displayName=$name;company=$company;mail=$_.mail;targetAddress=$targetAddress}; Break}
            default {@{givenName=$_.GivenName;sn=$_.sn;displayName=$name;company=$company;mail=$_.mail;targetAddress=$targetAddress}}
        }

        # Create contact in AD
        try { New-ADObject -Type contact -path $ENV:TARGETOU -name $Name -OtherAttributes $Params }
        catch { write-host $_ }
    }
} 
