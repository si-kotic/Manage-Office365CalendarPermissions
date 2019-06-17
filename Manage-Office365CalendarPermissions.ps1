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
        )][String]$AccessRights
    )
    Get-Mailbox $CalendarOwner | Foreach-Object {
        IF ($PSCmdlet.ShouldProcess($CalendarOwner)) {
            Add-MailboxFolderPermission "$_`:\Calendar" -User $User -AccessRights $AccessRights
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