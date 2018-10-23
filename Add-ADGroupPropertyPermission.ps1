function Add-ADGroupPropertyPermission {
    <#
    .SYNOPSIS
    This function is used for setting access rights on properties on Active Directory Groups.
    Use this code with caution! It has not been tested on a lot of objects/properties/access rights!
    .DESCRIPTION
    This function changes the ACLs on AD-Groups to enable granular delegation of them to other groups.
    Use this code with caution! It has not been tested on a lot of objects/properties/access rights!
    .EXAMPLE
    Add-ADGroupPropertyPermission -ADObject TheGroupSomeoneWantsAccessTo -MasterObject TheGroupWhoWillGainAccess -AccessRight WriteProperty -AccessRule Allow -Property Member -ActiveDirectoryServer MyDomain
    .PARAMETER ADObject
    Specify the identity of the group you want to delegate to the other group.
    .PARAMETER MasterObject
    Specify the identity of the group who should gain access to the specified property.
    .PARAMETER AccessRight
    Specify what access should be added, for example WriteProperty.
    .PARAMETER AccessRule
    Set this to Allow or Deny.
    .PARAMETER Property
    Specify which property this should be applied for.
    .PARAMETER ActiveDirectoryServer
    Specify domain or domain controller where the search for the groups will take place.
    #>

    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $True)]
        $ADObject,
        [Parameter(Mandatory = $True)]
        $MasterObject,
        [Parameter(Mandatory = $True)]
        $AccessRight,
        [Parameter(Mandatory = $True)]
        [ValidateSet("Allow", "Deny")]
        $AccessRule,
        [Parameter(Mandatory = $True)]
        $Property,
        $ActiveDirectoryServer = $(Get-ADDomain | Select-Object -ExpandProperty DNSRoot))

    # Load the AD objects
    try {
        $TheAccessGroup = Get-ADGroup -Identity $ADObject -Server $ActiveDirectoryServer -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to get the object with identity $ADObject. The error was: $($Error[0])."
        return
    }

    try {
        $TheOwnerGroup = Get-ADGroup -Identity $MasterObject -Server $ActiveDirectoryServer -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to get the object with identity $MasterObject. The error was: $($Error[0])."
        return
    }

    # Create SID-objects
    try {
        $OwnerGroupSid = New-Object System.Security.Principal.SecurityIdentifier ($TheOwnerGroup).SID -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to resolve the sid of $MasterObject. The error was: $($Error[0])."
        return
    }

    # Create the ACL object
    try {
        $AccessGroupACL = Get-Acl -Path "AD:\$($TheAccessGroup.DistinguishedName)" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to get the ACL of $($TheAccessGroup.DistinguishedName). The error was: $($Error[0])."
        return
    }

    #Get a reference to the RootDSE of the current domain
    $rootdse = Get-ADRootDSE

    #Create a hashtable to store the GUID value of the specified schema class and attribute
    $guidmap = @{}
    Get-ADObject -SearchBase ($rootdse.SchemaNamingContext) -LDAPFilter "(lDAPDisplayName=$Property)" -Properties lDAPDisplayName, schemaIDGUID | ForEach-Object {$guidmap[$_.lDAPDisplayName] = [System.GUID]$_.schemaIDGUID}

    # Allow time to create the object
    $AccessGroupACL.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $OwnerGroupSid, $AccessRight, $AccessRule, $guidmap["$Property"])) | Out-Null

    # Set the ACL
    try {
        Set-Acl -AclObject $AccessGroupACL -Path "AD:\$($TheAccessGroup.DistinguishedName)" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to set the ACL on $($TheAccessGroup.DistinguishedName). The error was: $($Error[0])."
        return
    }
}
