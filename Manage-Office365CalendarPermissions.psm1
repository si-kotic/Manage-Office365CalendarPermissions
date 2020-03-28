Function Add-Office365CalendarPermissions {
    [cmdletbinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory)][ValidateScript({
            Get-Mailbox $_
        })][String]$CalendarOwner,
        [Parameter(Mandatory)][ValidateScript({
            Get-Mailbox $_
        })][String]$User,
        [Parameter(Mandatory)][ValidateSet(
            "Author",
            "Contributor",
            "Editor",
            "None",
            "NonEditingAuthor",
            "Owner",
            "PublishingEditor",
            "PublishingAuthor",
            "Reviewer",
            "AvailabilityOnly",
            "LimitedDetails"
        )][String]$AccessRights,
        [ValidateSet(
            "None",
            "Delegate",
            "CanViewPrivateItems"
        )]$SharingPermissionFlags = "None"
    )
    Get-Mailbox $CalendarOwner | Foreach-Object {
        IF ($PSCmdlet.ShouldProcess($CalendarOwner)) {
            IF ($AccessRighta -eq "Editor") {
                Add-MailboxFolderPermission "$_`:\Calendar" -User $User -AccessRights $AccessRights -SharingPermissionFlags $SharingPermissionFlags -Confirm
            } ELSE {
                Add-MailboxFolderPermission "$_`:\Calendar" -User $User -AccessRights $AccessRights -Confirm
            }
        }
    }
}

Function Remove-Office365CalendarPermissions {
    [cmdletbinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory)][ValidateScript({
            Get-Mailbox $_
        })][String]$CalendarOwner,
        [Parameter(Mandatory)][ValidateScript({
            Get-Mailbox $_
        })][String]$User
    )
    Get-Mailbox $CalendarOwner | ForEach-Object {
        IF ($PSCmdlet.ShouldProcess($CalendarOwner)) {
            Remove-MailboxFolderPermission "$_`:\Calendar" -User $User
        }
    }
}

Function Get-Office365CalendarPermissions {
    Param (
        [Parameter(Mandatory)][ValidateScript({
            Get-Mailbox $_
        })][String]$CalendarOwner
    )
    Get-Mailbox $CalendarOwner | Foreach-Object {
        Get-MailboxFolderPermission "$_`:\Calendar"
    }
}