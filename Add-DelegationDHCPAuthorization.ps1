#####################################################################
# Add-DhcpAuthorizationDelegation.ps1
# Version 1.0
#
# Configures delegation in AD for DHCP authorization
#
# Ben Coremans
#
#####################################################################

function Add-DhcpAuthorizationDelegation {
    <#
.Synopsis
Configures delegation for DHCP authorization in Active Directory.
.Description
When delegating DHCP administration to an non Domain Administrator, 
you can use the build in Active Directory group DHCP Administrators to accomplish this task.
But during autorization of the DHCP server, you will not have the appropriate permissions. 
So authorization of the DHCP server require additional permissons in Active Directory. 
The delegation of authorization and unauthorization of DHCP servers is two-fold:
1. Granting permission to create/delete dHCPClass objects.
2. Granting permission to change all properties of the existing dHCPClass objects.
When this is done its is really possible to delegate DHCP administration.
.Parameter Identity
This function requires an AD object to grant the permissions to.
.EXAMPLE
Add-DhcpAuthorizationDelegation.ps1 "DHCP Authorization"

Running this command will granting permission for the identity "DHCP Authorization" 
to create/delete dHCPClass objects and to change all properties of the existing dHCPClass objects 
on the NetServices container in AD.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    $ADRootDSE = Get-ADRootDSE
    $ConfigNC = $ADRootDSE.configurationNamingContext
    $NetServicesPath = "AD:\CN=NetServices,CN=Services,$ConfigNC"
    $SchemaNamingContext = $ADRootDSE.SchemaNamingContext
    $guidmap = @{}
    Get-ADObject -SearchBase ($SchemaNamingContext) -LDAPFilter "(lDAPDisplayName=dHCPclass)" -Properties lDAPDisplayName, schemaIDGUID |
        Where-Object {$_.lDAPDisplayName -eq "dHCPclass"} | ForEach-Object {$guidmap[$_.lDAPDisplayName] = [System.GUID]$_.schemaIDGUID}
    $acl = Get-ACL $NetServicesPath 
    $InheritedObjectType = $guidmap['dHCPclass']

    Try {
        $account = New-Object System.Security.Principal.NTAccount($Identity)
        $sid = $account.Translate([System.Security.Principal.SecurityIdentifier])
    }
    Catch { Write-Error "There was an error while getting the SID for $Identity. Check if the identity exists. Error: $($Error[0].Exception.Message)"; return}

    # From here the new Access Rules will be created.
    # Allow permission to create/delete dHCPClass objects in the Container NetServices.
    $ActiveDirectoryRights = @("CreateChild", "DeleteChild")
    $AccessControlType = "Allow"
    $Inherit = "None"
    Try {
        $ACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, $ActiveDirectoryRights, $AccessControlType, $InheritedObjectType, $Inherit) -Verbose -ErrorAction Stop
        $acl.AddAccessRule($ace)
    }
    Catch { Write-Error "There was an error while adding the acces rule for $Identity. Error: $($Error[0].Exception.Message)"; return}

    # Allow permission to modify all properties of the existing dHCPClass objects.
    $ActiveDirectoryRights = @("WriteProperty", "ReadProperty", "ListChildren", "Delete")
    $AccessControlType = "Allow"
    $Inherit = "Children"

    Try { 
        $ACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, $ActiveDirectoryRights, $AccessControlType, $Inherit, $InheritedObjectType) -Verbose -ErrorAction Stop
        $acl.AddAccessRule($ace)
    }
    Catch { Write-Error "There was an error while adding the acces rule for $Identity. Error: $($Error[0].Exception.Message)"; return}

    # The accumulated access rules are set to NetServices AD container.
    Try {
        Set-ACL $NetServicesPath  -AclObject $acl -Verbose -ErrorAction Stop
    }
    Catch { Write-Error "There was an error while setting the new access rules for $Identity. Error: $($Error[0].Exception.Message)"}
}