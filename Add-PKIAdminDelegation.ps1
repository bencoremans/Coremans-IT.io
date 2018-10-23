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
    $SchemaNamingContext = $ADRootDSE.SchemaNamingContext

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
            $ACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, $AccessMask, $ObjAccessControlType, $ObjInherit) -Verbose #-ErrorAction Stop
            $acl.AddAccessRule($ace)
        }
        Catch { Write-Error "There was an error while adding the acces rule for $Identity. Error: $($Error[0].Exception.Message)"; return}

        Try {
            Set-ACL $AdContainerPath  -AclObject $acl -Verbose #-ErrorAction Stop
        }
        Catch { Write-Error "There was an error while setting the new access rules for $Identity. Error: $($Error[0].Exception.Message)"}
    }

    SetPermissions2Container -Identity $Identity -AdContainerPath "AD:\CN=Public Key Services,CN=Services,$ConfigNC" `
        -AccessMask @("GenericAll") -ObjAccessControlType "Allow" -ObjInherit "SelfAndChildren"


}
