#Check if module exists and install module with current user scope if it does not exist
write-host "Checking for AzureAD module" -ForegroundColor White -BackgroundColor Black

if (Get-installedModule -Name AzureAD -erroraction SilentlyContinue) {
    Write-Host "AzureAD Module exists" -ForegroundColor green -BackgroundColor Black
} 
else {
    write-host "AzureAD Module not installed, installing module with current user scope" -ForegroundColor yellow -BackgroundColor black
    install-module AzureAD -scope currentuser
}

#Connect to Azure AD with interactive prompt
write-host "Connecting to AzureAD" -ForegroundColor White -BackgroundColor Black
connect-AzureAD

#create a new Azure AD Application
write-host "Creating Azure AD Application" -ForegroundColor White -BackgroundColor Black
$appName = "ExpoE Exchange App"
$myApp = (Get-AzureADApplication -Filter "DisplayName eq '$($appName)'")

if($myApp -eq $null){
    $myApp = New-AzureADApplication -DisplayName $appName -ReplyUrls @("https://office.com/")
    write-host "App created sucessfully" -ForegroundColor green -BackgroundColor black
}
else {
    write-host "Application already exists" -ForegroundColor yellow -BackgroundColor black
} 

Start-Sleep -Seconds 2

#Add the current session user as the application owner
write-host "Assigning the application owner" -ForegroundColor white -BackgroundColor black
$currentUser = (Get-AzureADUser -ObjectId (Get-AzureADCurrentSessionInfo).Account.Id)
Add-AzureADApplicationOwner -ObjectId $myApp.ObjectId -RefObjectId $currentUser.ObjectId
write-host "Application owner granted sucessfully" -ForegroundColor green -BackgroundColor black

#
#
#Assign the required API permissions to the application
#
#

write-host "Assigning app permissions" -ForegroundColor white -BackgroundColor black
#Get Service Principal of Microsoft Graph Resource API 
$graphSP =  Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -eq "Office 365 Exchange Online"}
 
#Initialize RequiredResourceAccess for Microsoft Graph Resource API 
$requiredGraphAccess = New-Object Microsoft.Open.AzureAD.Model.RequiredResourceAccess
$requiredGraphAccess.ResourceAppId = $graphSP.AppId
$requiredGraphAccess.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]
 
#Set Application Permissions
$ApplicationPermissions = @('Exchange.ManageAsApp')
 
#Add app permissions
ForEach ($permission in $ApplicationPermissions) {
$reqPermission = $null
#Get required app permission
$reqPermission = $graphSP.AppRoles | Where-Object {$_.Value -eq $permission}
if($reqPermission)
{
$resourceAccess = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
$resourceAccess.Type = "Role"
$resourceAccess.Id = $reqPermission.Id    
#Add required app permission
$requiredGraphAccess.ResourceAccess.Add($resourceAccess)
write-host "App permissions found" -ForegroundColor green -BackgroundColor black
}
else
{
Write-Host "App permission $permission not found in the Graph Resource API" -ForegroundColor Red
}
}
 
#Set Delegated Permissions
#$DelegatedPermissions = @('') #Leave it as empty array if not required
 
#Add delegated permissions
#ForEach ($permission in $DelegatedPermissions) {
#$reqPermission = $null
#Get required delegated permission
#$reqPermission = $graphSP.Oauth2Permissions | Where-Object {$_.Value -eq $permission}
#if($reqPermission)
#{
#$resourceAccess = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
#$resourceAccess.Type = "Scope"
#$resourceAccess.Id = $reqPermission.Id    
#Add required delegated permission
#$requiredGraphAccess.ResourceAccess.Add($resourceAccess)
#}
#else
#{
#Write-Host "Delegated permission $permission not found in the Graph Resource API" -ForegroundColor Red
#}
#}
 
#Add required resource accesses
$requiredResourcesAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]
$requiredResourcesAccess.Add($requiredGraphAccess)
 
#Set permissions in existing Azure AD App
$appObjectId=$myApp.ObjectId
#$appObjectId="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
Set-AzureADApplication -ObjectId $appObjectId -RequiredResourceAccess $requiredResourcesAccess 
write-host "App permissions applied" -ForegroundColor green -BackgroundColor black
Start-Sleep -Seconds 2


#Grant consent in interactive session
$appid = $myApp.AppId
$TenantID = Get-AzureADTenantDetail
$tID = $TenantID.ObjectId

#Pause waiting for Azure
write-host "A new private window will open shortly. Please login with global admin and accept the application access. Once done, close the windows and click ENTER to continue" -ForegroundColor black -BackgroundColor yellow
Start-Sleep -Seconds 20

#Open edge in private windows and approve consent for app permissions
$url = "https://login.microsoftonline.com/" + $tID + "/adminconsent?client_id=" + $appid 
[System.Diagnostics.Process]::Start("msedge.exe", "-InPrivate $url")
	
read-host “Press ENTER to continue...”

Write-host "pausing for 20 seconds, please wait..."

Start-Sleep -Seconds 20

write-host "Assigning roles to service principal" -ForegroundColor white -BackgroundColor black
# Find Azure AD role by built in name
$role = Get-AzureADMSRoleDefinition -Filter "DisplayName eq 'Exchange Administrator'"

# Find Azure AD service principal by display name
$TeamsApp = Get-AzureADServicePrincipal -Filter "DisplayName eq 'ExpoE Exchange App'"

$roleid = $Role.id
$teamsappid = $TeamsApp.ObjectId
$assignedrole = Get-AzureADMSRoleAssignment -Filter "PrincipalId eq '$teamsappid' AND roleDefinitionId eq '$roleid'"

IF($assignedrole -eq $null){
New-AzureADMSRoleAssignment -RoleDefinitionId $role.Id -PrincipalId $TeamsApp.ObjectId -DirectoryScope "/"
write-host "Role assigned successfully" -ForegroundColor green -BackgroundColor black
}
Else {
write-host "Role is already assigned" -ForegroundColor yellow -BackgroundColor black
}

Connect-ExchangeOnline -CertificateThumbprint "8CD77A27525547D87C710A48633F81D9AD7C6AB6" -AppId "d921284b-a6ab-4328-a684-be9cf0510a05" -organization "expoetcaas.onmicrosoft.com" 
