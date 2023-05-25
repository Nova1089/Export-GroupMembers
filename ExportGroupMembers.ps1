<#
Allows you to search for an Office 365 group and export all members to CSV.
Does not work with Powershell 7.
#>

# Defining functions
Function UseModule([string]$moduleName)
{
    while ((IsModuleInstalled($moduleName)) -eq $false)
    {
        PromptToInstallModule($moduleName)
        TestSessionPrivileges
        Install-Module $moduleName
        if ((IsModuleInstalled($moduleName)) -eq $true)
        {
            Write-Host "Importing module..."
            Import-Module $moduleName
        }
        else
        {
            continue
        }
    }
}

Function IsModuleInstalled([string]$moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($module -ne $null)
}

Function PromptToInstallModule([string]$moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required."
        $confirmInstall = Read-Host -Prompt "Would you like to install it? (y/n)"
    }
    while ($confirmInstall -notmatch "\b[yY]\b")
}

Function TestSessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Throw "Please run script with admin privileges. 
            1. Open Powershell as admin.
            2. CD into script directory.
            3. Run .\scriptname.ps1"
    }
}

Function ConnectToOffice365
{
    Get-MsolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null

    while ($errorConnecting -ne $null)
    {
        Write-Host "Connecting to Office 365..."
        Connect-MsolService -ErrorAction SilentlyContinue
        Get-MSolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null   

        if ($errorConnecting -ne $null)
        {
            Read-Host -Prompt "Failed to connect to Office 365. Press Enter to try again."
        }
    }
}

Function PromptAndExport
{
    do
    {
        $group = PromptForGroup
        ExportGroupMembers($group)
        do
        {
            $goAgain = Read-Host -Prompt "Would you like to perform another export? (y/n)"
        }
        while ($goAgain -notmatch '\b[yYnN]\b')
    }
    while ($goAgain -match '\b[yY]\b')
}

Function PromptForGroup
{
    do
    {
        $groupPrompt = Read-Host -Prompt "Enter name of O365 group."
        $group = Get-MsolGroup -SearchString $groupPrompt

        if ($group.count -lt 1)
        {
            Write-Host "No group with that name was found."
        }
        elseif ($group.count -gt 1)
        {
            Write-Host "Found more than 1 group where name begins with that string."
        }
    }
    while ($group.count -ne 1)

    Write-Host "Group found. Display Name: $($group.DisplayName), Email: $($group.EmailAddress), Type: $($group.GroupType)"
    return $group
}

Function ExportGroupMembers([object]$group)
{
    Write-Host "Exporting to CSV..."
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $groupName = $group.DisplayName
    $timeStamp = NewTimeStamp    
    $path = "$desktopPath\$groupName Group Members $timeStamp.csv"
    Get-MsolGroupMember -All -GroupObjectId $group.ObjectId | Export-CSV -Path $path
    Write-Host "Finished exporting to $path."
}

Function NewTimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

# Main
UseModule("MSOnline")
ConnectToOffice365
PromptAndExport
Read-Host -Prompt "Press Enter to exit"