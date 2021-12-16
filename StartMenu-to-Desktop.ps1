<#
    script for creating desktop icons from StartMenu entries
    
    examples:
    
    # create a desktop icon for the installed app "Firefox"
    powershell.exe -executionpolicy bypass .\StartMenu-to-Desktop.ps1 -AppName "Firefox"
    
    # create a desktop icon for the installed app "Firefox" with the name "Firefox Webbrowser"
    powershell.exe -executionpolicy bypass .\StartMenu-to-Desktop.ps1 -AppName "Firefox" -IconName "Firefox Webbrowser" -Unique

    # create a desktop icon for the installed app "Firefox" and remove any other icons for the app on the desktop
    powershell.exe -executionpolicy bypass .\StartMenu-to-Desktop.ps1 -AppName "Firefox" -Unique
    
    # remove desktop icon named "Firefox"
    powershell.exe -executionpolicy bypass .\StartMenu-to-Desktop.ps1 -AppName "Firefox" -Remove
#>

param (
    [Parameter(Mandatory=$true)] # Name of the app for which a desktop icon is to be created
    [string]$AppName = '',

    [Parameter(Mandatory=$false)] # name for the desktop icon (optional)
    [string]$IconName = '',

    [Parameter(Mandatory=$false)] # if provided the desktop icon is removeed
    [switch]$Remove = $false,
    
    [Parameter(Mandatory=$false)] # if set, removes all other desktop icons for the program
    [switch]$Unique = $false
)

if (!$IconName) {
    $IconName = $AppName
}

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    write-host "SYSTEM context (Shortcut applies for all users)"
    $tagetPath = "$([Environment]::GetFolderPath('CommonDesktopDirectory'))"
}
else {
    write-host "USER context (Shortcut only applies for the logged on user)"
    $tagetPath = "$([Environment]::GetFolderPath("Desktop"))"
}

if ($Remove) {
    if (Test-Path("$tagetPath\$IconName.lnk")) {
        Write-Host "removing desktop icon for $AppName..."
        try {
            Remove-Item -Path "$tagetPath\$IconName.lnk" -Force
            exit 0
        }
        catch {
            Write-Output $_.Exception.Message
            exit 1
        }
    }
}
else {
    Write-Host "Trying to find App $AppName in StartMenu directories"
    $LNKfiles = ""
    @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:AppData\Microsoft\Windows\Start Menu\Programs"
    ) | foreach {
        $startMenuLNKs = Get-ChildItem -path $_ -recurse -Include *.lnk | Select Name,FullName | Sort Name -Descending   # most recent program version on the top
        if (!$LNKfiles) {
            $LNKfiles = $startMenuLNKs | Where-Object {$_.Name -like "$AppName.lnk"}
        }
        if (!$LNKfiles) {
            $LNKfiles = $startMenuLNKs | Where-Object {$_.Name -match "^$AppName [0-9.]+.lnk$"}
        }
        if ($IconName -notlike $AppName) {
            if (!$LNKfiles) {
                $LNKfiles = $startMenuLNKs | Where-Object {$_.Name -like "$IconName"}
            }
            if (!$LNKfiles) {
                $LNKfiles = $startMenuLNKs | Where-Object {$_.Name -match "^$IconName [0-9.]+.lnk$"}
            }
        }
    }
    if (!$LNKfiles) {
        Write-Output "Sorry, no matching LNK file found for App $AppName in start menu folders..."
        exit 1
    }
    else {
        $LNKfile = $LNKfiles | select -First 1
        Write-Output "Best match: $($LNKfile.FullName)"
        if ($Unique) {
            write-host "removing any existing desktop icons for the program"
            $currentDesktopIcons = Get-ChildItem -path "$tagetPath\*.lnk" -Force | Select Name,FullName | Sort Name
            $existingLNKfiles = $currentDesktopIcons | Where-Object {($_.Name -like "$AppName.lnk") -or ($_.Name -match "^$AppName [0-9.]+.lnk$") -or ($_.Name -like "$IconName") -or ($_.Name -match "^$IconName [0-9.]+.lnk$")}
            $existingLNKfiles | foreach {
                write-host "removing: $($_.Name)"
                Remove-Item -Path $_.FullName
            }
        }
        try {
            Copy-Item -Path $LNKfile.FullName -Destination "$tagetPath\$IconName.lnk"
            exit 0
        }
        catch {
            Write-Output $_.Exception.Message
            exit 1
        }
    }
}
