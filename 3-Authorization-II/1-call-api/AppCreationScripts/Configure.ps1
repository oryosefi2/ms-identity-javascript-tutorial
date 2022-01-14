
[CmdletBinding()]
param(
    [Parameter(Mandatory=$False, HelpMessage='Tenant ID (This is a GUID which represents the "Directory ID" of the AzureAD tenant into which you want to create the apps')]
    [string] $tenantId,
    [Parameter(Mandatory=$False, HelpMessage='Azure environment to use while running the script. Default = Global')]
    [string] $azureEnvironmentName
)

<#
 This script creates the Azure AD applications needed for this sample and updates the configuration files
 for the visual Studio projects from the data in the Azure AD applications.

 In case you don't have Microsoft.Graph.Applications already installed, the script will automatically install it for the current user
 
 There are four ways to run this script. For more information, read the AppCreationScripts.md file in the same folder as this script.
#>

# Adds the requiredAccesses (expressed as a pipe separated string) to the requiredAccess structure
# The exposed permissions are in the $exposedPermissions collection, and the type of permission (Scope | Role) is 
# described in $permissionType
Function AddResourcePermission($requiredAccess, `
                               $exposedPermissions, [string]$requiredAccesses, [string]$permissionType)
{
    foreach($permission in $requiredAccesses.Trim().Split("|"))
    {
        foreach($exposedPermission in $exposedPermissions)
        {
            if ($exposedPermission.Value -eq $permission)
                {
                $resourceAccess = New-Object Microsoft.Graph.PowerShell.Models.MicrosoftGraphResourceAccess
                $resourceAccess.Type = $permissionType # Scope = Delegated permissions | Role = Application permissions
                $resourceAccess.Id = $exposedPermission.Id # Read directory data
                $requiredAccess.ResourceAccess += $resourceAccess
                }
        }
    }
}

#
# Example: GetRequiredPermissions "Microsoft Graph"  "Graph.Read|User.Read"
# See also: http://stackoverflow.com/questions/42164581/how-to-configure-a-new-azure-ad-application-through-powershell
Function GetRequiredPermissions([string] $applicationDisplayName, [string] $requiredDelegatedPermissions, [string]$requiredApplicationPermissions, $servicePrincipal)
{
    # If we are passed the service principal we use it directly, otherwise we find it from the display name (which might not be unique)
    if ($servicePrincipal)
    {
        $sp = $servicePrincipal
    }
    else
    {
        $sp = Get-MgServicePrincipal -Filter "DisplayName eq '$applicationDisplayName'"
    }
    $appid = $sp.AppId
    $requiredAccess = New-Object Microsoft.Graph.PowerShell.Models.MicrosoftGraphRequiredResourceAccess
    $requiredAccess.ResourceAppId = $appid 
    $requiredAccess.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Graph.PowerShell.Models.MicrosoftGraphResourceAccess]

    # $sp.Oauth2Permissions | Select Id,AdminConsentDisplayName,Value: To see the list of all the Delegated permissions for the application:
    if ($requiredDelegatedPermissions)
    {
        AddResourcePermission $requiredAccess -exposedPermissions $sp.Oauth2PermissionScopes -requiredAccesses $requiredDelegatedPermissions -permissionType "Scope"
    }
    
    # $sp.AppRoles | Select Id,AdminConsentDisplayName,Value: To see the list of all the Application permissions for the application
    if ($requiredApplicationPermissions)
    {
        AddResourcePermission $requiredAccess -exposedPermissions $sp.AppRoles -requiredAccesses $requiredApplicationPermissions -permissionType "Role"
    }
    return $requiredAccess
}


Function UpdateLine([string] $line, [string] $value)
{
    $index = $line.IndexOf(':')
    $lineEnd = ''

    if($line[$line.Length - 1] -eq ','){   $lineEnd = ',' }
    
    if ($index -ige 0)
    {
        $line = $line.Substring(0, $index+1) + " " + '"' + $value+ '"' + $lineEnd
    }
    return $line
}

Function UpdateTextFile([string] $configFilePath, [System.Collections.HashTable] $dictionary)
{
    $lines = Get-Content $configFilePath
    $index = 0
    while($index -lt $lines.Length)
    {
        $line = $lines[$index]
        foreach($key in $dictionary.Keys)
        {
            if ($line.Contains($key))
            {
                $lines[$index] = UpdateLine $line $dictionary[$key]
            }
        }
        $index++
    }

    Set-Content -Path $configFilePath -Value $lines -Force
}

Function ReplaceInLine([string] $line, [string] $key, [string] $value)
{
    $index = $line.IndexOf($key)
    if ($index -ige 0)
    {
        $index2 = $index+$key.Length
        $line = $line.Substring(0, $index) + $value + $line.Substring($index2)
    }
    return $line
}

Function ReplaceInTextFile([string] $configFilePath, [System.Collections.HashTable] $dictionary)
{
    $lines = Get-Content $configFilePath
    $index = 0
    while($index -lt $lines.Length)
    {
        $line = $lines[$index]
        foreach($key in $dictionary.Keys)
        {
            if ($line.Contains($key))
            {
                $lines[$index] = ReplaceInLine $line $key $dictionary[$key]
            }
        }
        $index++
    }

    Set-Content -Path $configFilePath -Value $lines -Force
}
<#.Description
   This function creates a new Azure AD scope (OAuth2Permission) with default and provided values
#>  
Function CreateScope( [string] $value, [string] $userConsentDisplayName, [string] $userConsentDescription, [string] $adminConsentDisplayName, [string] $adminConsentDescription)
{
    $scope = New-Object Microsoft.Graph.PowerShell.Models.MicrosoftGraphPermissionScope
    $scope.Id = New-Guid
    $scope.Value = $value
    $scope.UserConsentDisplayName = $userConsentDisplayName
    $scope.UserConsentDescription = $userConsentDescription
    $scope.AdminConsentDisplayName = $adminConsentDisplayName
    $scope.AdminConsentDescription = $adminConsentDescription
    $scope.IsEnabled = $true
    $scope.Type = "User"
    return $scope
}

<#.Description
   This function creates a new Azure AD AppRole with default and provided values
#>  
Function CreateAppRole([string] $types, [string] $name, [string] $description)
{
    $appRole = New-Object Microsoft.Graph.PowerShell.Models.MicrosoftGraphAppRole
    $appRole.AllowedMemberTypes = New-Object System.Collections.Generic.List[string]
    $typesArr = $types.Split(',')
    foreach($type in $typesArr)
    {
        $appRole.AllowedMemberTypes += $type;
    }
    $appRole.DisplayName = $name
    $appRole.Id = New-Guid
    $appRole.IsEnabled = $true
    $appRole.Description = $description
    $appRole.Value = $name;
    return $appRole
}

Function ConfigureApplications
{
    <#.Description
       This function creates the Azure AD applications for the sample in the provided Azure AD tenant and updates the
       configuration files in the client and service project  of the visual studio solution (App.Config and Web.Config)
       so that they are consistent with the Applications parameters
    #> 
    
    if (!$azureEnvironmentName)
    {
        $azureEnvironmentName = "Global"
    }

    # Connect to the Microsoft Graph API, non-interactive is not supported for the moment (Oct 2021)
    Write-Host "Connecting Microsoft Graph"
    if ($tenantId -eq "") {
        Connect-MgGraph -Scopes "Application.ReadWrite.All" -Environment $azureEnvironmentName
        $tenantId = (Get-MgContext).TenantId
    }
    else {
        Connect-MgGraph -TenantId $tenantId -Scopes "Application.ReadWrite.All" -Environment $azureEnvironmentName
    }
    

   # Create the spa AAD application
   Write-Host "Creating the AAD application (ms-identity-javascript-tutorial-c3s1-spa)"
   
   # create the application 
   $spaAadApplication = New-MgApplication -DisplayName "ms-identity-javascript-tutorial-c3s1-spa" `
                                                   -Spa `
                                                   @{ `
                                                       RedirectUris = "http://localhost:3000/"; `
                                                    } `
                                                   -SignInAudience AzureADMyOrg `
                                                   #end of command
    $tenantName = (Get-MgApplication -ApplicationId $spaAadApplication.Id).PublisherDomain
    Update-MgApplication -ApplicationId $spaAadApplication.Id -IdentifierUris @("https://$tenantName/ms-identity-javascript-tutorial-c3s1-spa")
    
    # create the service principal of the newly created application 
    $currentAppId = $spaAadApplication.AppId
    $spaServicePrincipal = New-MgServicePrincipal -AppId $currentAppId -Tags {WindowsAzureActiveDirectoryIntegratedApp}

    # add the user running the script as an app owner if needed
    $owner = Get-MgApplicationOwner -ApplicationId $spaAadApplication.Id
    if ($owner -eq $null)
    { 
        New-MgApplicationOwnerByRef -ApplicationId $spaAadApplication.Id  -BodyParameter = @{"@odata.id" = "htps://graph.microsoft.com/v1.0/directoryObjects/$user.ObjectId"}
        Write-Host "'$($user.UserPrincipalName)' added as an application owner to app '$($spaServicePrincipal.DisplayName)'"
    }
    Write-Host "Done creating the spa application (ms-identity-javascript-tutorial-c3s1-spa)"

    # URL of the AAD application in the Azure portal
    # Future? $spaPortalUrl = "https://portal.azure.com/#@"+$tenantName+"/blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/"+$spaAadApplication.AppId+"/objectId/"+$spaAadApplication.ObjectId+"/isMSAApp/"
    $spaPortalUrl = "https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/"+$spaAadApplication.AppId+"/objectId/"+$spaAadApplication.ObjectId+"/isMSAApp/"
    Add-Content -Value "<tr><td>spa</td><td>$currentAppId</td><td><a href='$spaPortalUrl'>ms-identity-javascript-tutorial-c3s1-spa</a></td></tr>" -Path createdApps.html
    $requiredResourcesAccess = New-Object System.Collections.Generic.List[Microsoft.Graph.PowerShell.Models.MicrosoftGraphRequiredResourceAccess]

    
    # Add Required Resources Access (from 'spa' to 'Microsoft Graph')
    Write-Host "Getting access from 'spa' to 'Microsoft Graph'"
    $requiredPermissions = GetRequiredPermissions -applicationDisplayName "Microsoft Graph" `
        -requiredDelegatedPermissions "User.Read" `
    

    $requiredResourcesAccess.Add($requiredPermissions)
    Update-MgApplication -ApplicationId $spaAadApplication.Id -RequiredResourceAccess $requiredResourcesAccess
    Write-Host "Granted permissions."

   # Create the service AAD application
   Write-Host "Creating the AAD application (ms-identity-javascript-tutorial-c3s1-api)"
   
   # create the application 
   $serviceAadApplication = New-MgApplication -DisplayName "ms-identity-javascript-tutorial-c3s1-api" `
                                                       -Web `
                                                       @{ `
                                                           HomePageUrl = "http://localhost:5000/api"; `

                                                           ImplicitGrantSettings = @{EnableIdTokenIssuance=$true;EnableAccessTokenIssuance=$false} ` 
 
                                                         } `
                                                       -SignInAudience AzureADMyOrg `
                                                       #end of command
    $serviceIdentifierUri = 'api://'+$serviceAadApplication.AppId
    Update-MgApplication -ApplicationId $serviceAadApplication.Id -IdentifierUris @($serviceIdentifierUri)
    
    # create the service principal of the newly created application 
    $currentAppId = $serviceAadApplication.AppId
    $serviceServicePrincipal = New-MgServicePrincipal -AppId $currentAppId -Tags {WindowsAzureActiveDirectoryIntegratedApp}

    # add the user running the script as an app owner if needed
    $owner = Get-MgApplicationOwner -ApplicationId $serviceAadApplication.Id
    if ($owner -eq $null)
    { 
        New-MgApplicationOwnerByRef -ApplicationId $serviceAadApplication.Id  -BodyParameter = @{"@odata.id" = "htps://graph.microsoft.com/v1.0/directoryObjects/$user.ObjectId"}
        Write-Host "'$($user.UserPrincipalName)' added as an application owner to app '$($serviceServicePrincipal.DisplayName)'"
    }
    
    # rename the user_impersonation scope if it exists to match the readme steps or add a new scope
       
    # delete default scope i.e. User_impersonation
    $scope = $serviceAadApplication.Api.Oauth2PermissionScopes | Where-Object { $_.Value -eq "User_impersonation" }
    if($scope -ne $null)
    {
        # disable the scope
        $scope.IsEnabled = $false
        Update-MgApplication -ApplicationId $serviceAadApplication.Id -Api @{Oauth2PermissionScopes = $($scope)}

        # clear the scope
        Update-MgApplication -ApplicationId $serviceAadApplication.Id -Api @{Oauth2PermissionScopes = $()}
    }
    $scopes = New-Object System.Collections.Generic.List[Microsoft.Graph.PowerShell.Models.MicrosoftGraphPermissionScope]
    $scope = CreateScope -value access_as_user  `
    -userConsentDisplayName "Access ms-identity-javascript-tutorial-c3s1-api"  `
    -userConsentDescription "Allow the application to access ms-identity-javascript-tutorial-c3s1-api on your behalf."  `
    -adminConsentDisplayName "Access ms-identity-javascript-tutorial-c3s1-api"  `
    -adminConsentDescription "Allows the app to have the same access to information in the directory on behalf of the signed-in user."
            
    $scopes.Add($scope)
    
    # add/update scopes
    Update-MgApplication -ApplicationId $serviceAadApplication.Id -Api @{Oauth2PermissionScopes = $scopes}
    Write-Host "Done creating the service application (ms-identity-javascript-tutorial-c3s1-api)"

    # URL of the AAD application in the Azure portal
    # Future? $servicePortalUrl = "https://portal.azure.com/#@"+$tenantName+"/blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/"+$serviceAadApplication.AppId+"/objectId/"+$serviceAadApplication.ObjectId+"/isMSAApp/"
    $servicePortalUrl = "https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/"+$serviceAadApplication.AppId+"/objectId/"+$serviceAadApplication.ObjectId+"/isMSAApp/"
    Add-Content -Value "<tr><td>service</td><td>$currentAppId</td><td><a href='$servicePortalUrl'>ms-identity-javascript-tutorial-c3s1-api</a></td></tr>" -Path createdApps.html
    
    # Update config file for 'spa'
    $configFile = $pwd.Path + "\..\SPA\App\authConfig.js"
    $dictionary = @{ "Enter_the_Application_Id_Here" = $spaAadApplication.AppId;"Enter_the_Tenant_Info_Here" = $tenantId;"Enter_the_Redirect_Uri_Here" = $spaAadApplication.Spa.RedirectUris;"Enter_the_Web_Api_Uri_Here" = $serviceAadApplication.Web.HomePageUrl;"Enter_the_Web_Api_Scope_Here" = ("api://"+$serviceAadApplication.AppId+"/access_as_user") };

    Write-Host "Updating the sample code ($configFile)"

    ReplaceInTextFile -configFilePath $configFile -dictionary $dictionary
    
    # Update config file for 'service'
    $configFile = $pwd.Path + "\..\API\config.json"
    $dictionary = @{ "clientID" = $serviceAadApplication.AppId;"tenantID" = $tenantId;"audience" = $serviceAadApplication.AppId };

    Write-Host "Updating the sample code ($configFile)"

    UpdateTextFile -configFilePath $configFile -dictionary $dictionary
    Write-Host -ForegroundColor Green "------------------------------------------------------------------------------------------------" 
    Write-Host "IMPORTANT: Please follow the instructions below to complete a few manual step(s) in the Azure portal":
    Write-Host "- For service"
    Write-Host "  - Navigate to $servicePortalUrl"
    Write-Host "  - Navigate to the Manifest page, find the property 'accessTokenAcceptedVersion' and set it to '2'" -ForegroundColor Red 
    Write-Host -ForegroundColor Green "------------------------------------------------------------------------------------------------" 
       if($isOpenSSL -eq 'Y')
    {
        Write-Host -ForegroundColor Green "------------------------------------------------------------------------------------------------" 
        Write-Host "You have generated certificate using OpenSSL so follow below steps: "
        Write-Host "Install the certificate on your system from current folder."
        Write-Host -ForegroundColor Green "------------------------------------------------------------------------------------------------" 
    }
    Add-Content -Value "</tbody></table></body></html>" -Path createdApps.html  
}

# Pre-requisites
if ($null -eq (Get-Module -ListAvailable -Name "Microsoft.Graph.Applications")) {
    Install-Module "Microsoft.Graph.Applications" -Scope CurrentUser 
}

Import-Module Microsoft.Graph.Applications

Set-Content -Value "<html><body><table>" -Path createdApps.html
Add-Content -Value "<thead><tr><th>Application</th><th>AppId</th><th>Url in the Azure portal</th></tr></thead><tbody>" -Path createdApps.html

$ErrorActionPreference = "Stop"

# Run interactively (will ask you for the tenant ID)
ConfigureApplications -tenantId $tenantId -environment $azureEnvironmentName

Write-Host "Disconnecting from tenant"
Disconnect-MgGraph