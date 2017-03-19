Function Add-O365MailboxFolderPermissions {
#requires -version 2
[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
    
	[Parameter( Mandatory=$true)]
	[string]$User,

    [Parameter( Mandatory=$true)]
	[ValidateSet("Author","Contributor","Editor","NonEditingAuthor","Owner","PublishingEditor","PublishingAuthor","Reviewer")]
    [string]$PermissionType
)

#...................................
# Variables
#...................................
$exclusions = @("/Sync Issues",
                "/Sync Issues/Conflicts",
                "/Sync Issues/Local Failures",
                "/Sync Issues/Server Failures",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions"
                )


#...................................
# Script
#...................................

$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

foreach ($mailboxfolder in $mailboxfolders)
{
    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace(“\Top of Information Store”,”\”)
    }
    $identity = "$($mailbox):$folder"
    Write-Host "Adding $user to $identity with $access permissions"
    try
    {
        Add-MailboxFolderPermission "$identity" -User $user -Access $PermissionType -ErrorAction STOP
    }
    catch
    {
        Write-Warning $_.Exception.Message
    }
}


#...................................
# End
#...................................
}
function Remove-O365MailboxFolderPermissions{
#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
    
	[Parameter( Mandatory=$true)]
	[string]$User
)


#...................................
# Variables
#...................................

$exclusions = @("/Sync Issues",
                "/Sync Issues/Conflicts",
                "/Sync Issues/Local Failures",
                "/Sync Issues/Server Failures",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions"
                )


#...................................
# Initialize
#...................................

#...................................
# Script
#...................................

    $mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

    foreach ($mailboxfolder in $mailboxfolders)
                                                                                {
    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace(“\Top of Information Store”,”\”)
    }
    $identity = "$($mailbox):$folder"
    Write-Host "Checking $identity for permissions for user $user"
    if (Get-MailboxFolderPermission -Identity $identity -User $user -ErrorAction SilentlyContinue)
    {
        try
        {
            Remove-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -ErrorAction STOP
            Write-Host -ForegroundColor Green "Removed!"
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }
    }
    }


#...................................
# End
#...................................

}
