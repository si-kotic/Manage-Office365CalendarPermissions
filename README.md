# Manage-Office365CalendarPermissions
This module contains functions which allow you to retrieve, add and remove permissiosn to the Calendar of an Office 365 Mailbox.

## Cmdlets
* Add-Office365CalendarPermissions
* Remove-Office365CalendarPermissions
* Get-Office365CalendarPermissions

## Usage
To install the module, download the files from GitHub and save them to your local device, eg `C:\PowerShellModules\Manage-Office365CalendarPermissions\`.  Next run the Import-Module command, providing the full path of the `.psm1` file to the `Name` parameter, eg `Import-Module -Name C:\PowerShellModules\Manage-Office365CalendarPermissions\Manage-Office365CalendarPermissions.psm1`.
### Add-Office365CalendarPermissions
The Add-Office365CalendarPermissions cmdlet allows you to grant one user specific permissions over another users Office 365 Calendar.  Confirmation is requested before any changes are applied.
#### Parameters
##### CalendarOwner
Specifies the Calendar to which you wish to grant permissions.  Accepts `string` values in the form of `user.name` or `User Name`.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### User
Specifies the User for whom you wish to grant permissions.  Accepts `string` values in the form of `user.name` or `User Name`.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### Access Rights
Specifies the permissions you wish to grant.  Accepts one of the following values: "Author", "Contributor", "Editor", "None", "NonEditingAuthor", "Owner", "PublishingEditor", "PublishingAuthor", "Reviewer", "AvailabilityOnly", "LimitedDetails".

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
#### Syntax
```powershell
Add-Office365CalendarPermissions -CalendarOwner "Big Boss Man" -User "Personal Assistant" -AccessRights Editor
```
#### Example
```
C:\>Add-Office365CalendarPermissions -CalendarOwner "Big Boss Man" -User "Personal Assistant" -AccessRights Editor

Confirm
Are you sure you want to perform this action?
Adding mailbox folder permission on "Big Boss Man:\Calendar" for user "Personal Assistant", access rights "'Editor'".
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [?] Help (default is "Y"): y
```
### Remove-Office365CalendarPermissions
The Remove-Office365CalendarPermissions cmdlet allows you to remove a users permissions over another users Office 365 Calendar.  Confirmation is requested before any changes are applied.
#### Parameters
##### CalendarOwner
Specifies the Calendar to which you wish to remove permissions.  Accepts `string` values in the form of `user.name` or `User Name`.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
##### User
Specifies the User for whom you wish to remove permissions.  Accepts `string` values in the form of `user.name` or `User Name`.

Argument | Value
--- | ---
Type | String
Position | Named
Default value | None
Accept pipeline input | False
Accept wildcard characters | False
Mandatory | True
#### Syntax
```powershell
Remove-Office365CalendarPermissions -CalendarOwner "Big Boss Man" -User "Personal Assistant"
```
#### Example
```
C:\>Remove-Office365CalendarPermissions -CalendarOwner "Big Boss Man" -User "Personal Assistant"

Confirm
Are you sure you want to perform this action?
Removing mailbox folder permission on "Big Boss Man:\Calendar" for user "Personal Assistant".
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [?] Help (default is "Y"): y
```
### Get-Office365CalendarPermissions
The Get-Office365CalendarPermissions cmdlet allows you to view who has which permissions over a specific users Office 365 Calendar.
#### Parameters
There are no parameters for this cmdlet.
#### Syntax
```powershell
Get-Office365CalendarPermissions -CalendarOwner "Big Boss Man"
```
#### Example
```
C:\>Get-Office365CalendarPermissions -CalendarOwner "Big Boss Man"
FolderName           User                 AccessRights                              SharingPermissionFlags
Calendar             Default              {Reviewer}
Calendar             Anonymous            {None}
Calendar             Personal Assistant   {Editor}
```