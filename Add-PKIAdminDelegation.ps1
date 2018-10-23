#####################################################################
# Add-PKIAdminDelegation.ps1
# Version 1.0
#
# Configures delegation in AD for PKI administration
#
# Ben Coremans
#
#####################################################################

function Add-PKIAdminDelegation {
    <#
.Synopsis
    Configures delegation for PKI Administration in Active Directory.
.Description
    When delegating PKI administration to an non Domain Administrator, 
    you need to delegate the required additional permissons in Active Directory. 
    The delegation of authorization and unauthorization of PKI servers is as follows:
    Granting full control permission to all objects and all descendants in the Container Public Key Services.
.Parameter Identity
    This function requires an AD object to grant the permissions to.
.EXAMPLE
    Add-PKIAdminDelegation.ps1 "PKI Admins"
    
    Running this command will granting full control permission for the identity "PKI Admins" 
    to all objects and all descendants in the Container Public Key Services in AD.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    $ADRootDSE = Get-ADRootDSE
    $ConfigNC = $ADRootDSE.configurationNamingContext
    Function SetPermissions2Container {
        [CmdletBinding()]
        param(
            [Parameter(Position = 0, mandatory = $true)]
            [string]$Identity,
            [Parameter(Position = 1, mandatory = $true)]
            [string]$AdContainerPath,
            [Parameter(Position = 2, mandatory = $true)]
            [array]$AccessMask,
            [Parameter(Position = 3, mandatory = $true)]
            [string]$ObjAccessControlType,
            [Parameter(Position = 4, mandatory = $true)]
            [string]$ObjInherit
        )

        Try {
            $account = New-Object System.Security.Principal.NTAccount($Identity)
            $sid = $account.Translate([System.Security.Principal.SecurityIdentifier])
        }
        Catch { Write-Error "There was an error while getting the SID for $Identity. Check if the identity exists. Error: $($Error[0].Exception.Message)"; return}
    
        $acl = Get-ACL $AdContainerPath 
        # From here the new Access Rules will be created.
        # Allow full control permission to all objects and all descendants in the Container Public Key Services.

        Try {
            #$ACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, $AccessMask, $ObjAccessControlType, $ObjInherit) -Verbose #-ErrorAction Stop
            #([GUID]("00000000-0000-0000-0000-000000000000")).guid
            $ACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, $AccessMask, $ObjAccessControlType, ([GUID]("00000000-0000-0000-0000-000000000000")).guid, $ObjInherit, ([GUID]("00000000-0000-0000-0000-000000000000")).guid)
            $acl.AddAccessRule($ace)
        }
        Catch { Write-Error "There was an error while adding the acces rule for $Identity. Error: $($Error[0].Exception.Message)"; return}

        Try {
            Set-ACL $AdContainerPath  -AclObject $acl -Verbose #-ErrorAction Stop
        }
        Catch { Write-Error "There was an error while setting the new access rules for $Identity. Error: $($Error[0].Exception.Message)"; return}
    }


    function Add-ADGroupPropertyPermission {
        <#
    .SYNOPSIS
    This function is used for setting access rights on properties on Active Directory Groups.
    Use this code with caution! It has not been tested on a lot of objects/properties/access rights!
    .DESCRIPTION
    This function changes the ACLs on AD-Groups to enable granular delegation of them to other groups.
    Use this code with caution! It has not been tested on a lot of objects/properties/access rights!
    .EXAMPLE
    Add-ADGroupPropertyPermission -Identity TheGroupWhoWillGainAccess -MasterObject TheGroupSomeoneWantsAccessTo -AccessRight WriteProperty -AccessRule Allow -Property Member -ActiveDirectoryServer MyDomain
    .PARAMETER Identity
    Specify the identity of the group who should gain access to the specified property.    
    .PARAMETER MasterObject
    Specify the identity of the group you want to delegate to the other group.
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
            $Identity,
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
            $TheAccessGroup = Get-ADGroup -Identity $MasterObject -Server $ActiveDirectoryServer -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to get the object with identity $MasterObject. The error was: $($Error[0])."
            return
        }

        try {
            $TheOwnerGroup = Get-ADGroup -Identity $Identity -Server $ActiveDirectoryServer -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to get the object with identity $Identity. The error was: $($Error[0])."
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
            Set-Acl -AclObject $AccessGroupACL -Path "AD:\$($TheAccessGroup.DistinguishedName)" -Verbose -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to set the ACL on $($TheAccessGroup.DistinguishedName). The error was: $($Error[0])."
            return
        }
    }

    SetPermissions2Container -Identity $Identity -AdContainerPath "AD:\CN=Public Key Services,CN=Services,$ConfigNC" -AccessMask @("GenericAll") -ObjAccessControlType "Allow" -ObjInherit "SelfAndChildren"
    Add-ADGroupPropertyPermission -Identity $Identity -MasterObject "Cert Publishers" -AccessRight WriteProperty -AccessRule Allow -Property member
    Add-ADGroupPropertyPermission -Identity $Identity -MasterObject "Pre-Windows 2000 Compatible Access" -AccessRight WriteProperty -AccessRule Allow -Property member
}
