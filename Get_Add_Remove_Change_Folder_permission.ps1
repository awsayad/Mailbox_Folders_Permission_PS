<#
.SYNOPSIS 
Use this script to Get Mailbox folder permissions, Add / Remove / Change permissions as needed.
 
.DESCRIPTION  
This script allows you to check the folders' permissions of a specific mailbox and change permissions accordingly. You can select an action
from the menu as needed.
 
.OUTPUTS 
Results are printed to the console.
 
.NOTES 
Written by: Aws Ayad

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS
CODE REMAINS WITH THE USER.
#>

$Loop = $true
While ($Loop)
{
write-host 
write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
write-host   'Mailbox Folder Permission Action Menu'
write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
write-host
write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
write-host -ForegroundColor white  -BackgroundColor DarkGreen   'Action Menu:' 
write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
write-host                                              ' 1)  Get mailbox folders list'
write-host                                              ' 2)  Get Access Rights of all folders'
write-host                                              ' 3)  Add explicit permission / Access Rights on all folders'
write-host                                              ' 4)  Removing user permission / Access Rights from all folders'
write-host                                              ' 5)  Change user current permission / Access Rights on all'
write-host                                              ' 6)  Exit Powershell Script'
write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
write-host -ForegroundColor white  -BackgroundColor DarkRed 'End of Action menu' 
write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
write-host

$menu = Read-Host "Select an option [1-6]"
write-host $menu

switch ($menu)

{
    
1
{
    try 
    {
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    $MailboxOwner = Read-Host "Enter user email address"

    write-host "Getting the mailbox folders list" -ForegroundColor yellow -BackgroundColor DarkGreen
    write-host "Please wait....." -ForegroundColor yellow -BackgroundColor DarkGreen

    Get-MailboxFolderStatistics $MailboxOwner | Select folderpath
    }

finally
    {
    write-host "*****************************************************************************************************"
    write-host "Done, script completed successfully" -ForegroundColor white -BackgroundColor Red
    write-host "*****************************************************************************************************"
    write-host "`n"       
    }
}

2
{
    try 
    {
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    $MailboxOwner = Read-Host "Enter user email address"

    write-host "Get current folders Access Rights (all folders)" -ForegroundColor yellow -BackgroundColor DarkGreen
    write-host "Please wait....." -ForegroundColor yellow -BackgroundColor DarkGreen

    foreach ($Folder in (Get-MailboxFolderStatistics $MailboxOwner )) `
    {get-MailboxFolderPermission -Identity "$($MailboxOwner):$($Folder.FolderPath.ToString().Replace("/","\").Replace([char]63743,"/"))" `
    -ErrorAction Ignore}
    }

finally
    {
    write-host "*****************************************************************************************************"
    write-host "Done, script completed successfully" -ForegroundColor white -BackgroundColor Red
    write-host "*****************************************************************************************************"
    write-host "`n"       
    }
}

3
{
    try 
    {
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    $MailboxOwner = Read-Host "Owner mailbox email address"
    $User = Read-Host "User email address"
    $AccessRights =  Read-Host "Please enter Access Rights (i.e. Reviewer, Editor, owner...etc.)"

    write-host "Add explicit permission / Access Rights on all folders" -ForegroundColor yellow -BackgroundColor DarkGreen
    write-host "Please wait....." -ForegroundColor yellow -BackgroundColor DarkGreen

    foreach ($Folder in (Get-MailboxFolderStatistics $MailboxOwner )) `
    {Add-MailboxFolderPermission -Identity "$($MailboxOwner):$($Folder.FolderPath.ToString().Replace("/","\").Replace([char]63743,"/"))" `
    -User $User -AccessRights $AccessRights -Confirm:$false -ErrorAction Ignore}
    }

finally
    {
    write-host "*****************************************************************************************************"
    write-host "Done, script completed successfully" -ForegroundColor white -BackgroundColor Red
    write-host "*****************************************************************************************************"
    write-host "`n"       
    }

}

4
{
    try 
    {
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    $MailboxOwner = Read-Host "Owner Mailbox email address"
    $User = Read-Host "User email address"

    write-host "Removing user permission / Access Rights from all folders" -ForegroundColor yellow -BackgroundColor DarkGreen
    write-host "Please wait....." -ForegroundColor yellow -BackgroundColor DarkGreen

    foreach ($Folder in (Get-MailboxFolderStatistics $MailboxOwner )) `
    {Remove-MailboxFolderPermission -Identity "$($MailboxOwner):$($Folder.FolderPath.ToString().Replace("/","\").Replace([char]63743,"/"))" `
    -User $User -Confirm:$false -ErrorAction Ignore}
    }

finally
    {
    write-host "*****************************************************************************************************"
    write-host "Done, script completed successfully" -ForegroundColor white -BackgroundColor Red
    write-host "*****************************************************************************************************"
    write-host "`n"       
    }
}

5
{
    try 
    {
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"
    
    $MailboxOwner = Read-Host "Owner Mailbox email address"
    $User = Read-Host "User email address"
    $AccessRights =  Read-Host "Please enter required Access Rights (i.e. Reviewer, Editor, owner...etc.)"

    write-host "Changing user explicit permission / Access Rights" -ForegroundColor yellow -BackgroundColor DarkGreen
    write-host "Please wait....." -ForegroundColor yellow -BackgroundColor DarkGreen
    
    foreach ($Folder in (Get-MailboxFolderStatistics $MailboxOwner )) `
    {set-MailboxFolderPermission -Identity "$($MailboxOwner):$($Folder.FolderPath.ToString().Replace("/","\").Replace([char]63743,"/"))" `
    -User $User -AccessRights $AccessRights -Confirm:$false -ErrorAction Ignore}
    }

    finally
    {
        write-host "*****************************************************************************************************"
        write-host "Done, script completed successfully" -ForegroundColor white -BackgroundColor Red
        write-host "*****************************************************************************************************"
        write-host "`n"       
    }

}

6
{
$Loop = $true
Exit
}

}
}