<#
.SYNOPSIS
Script in PowerShell to create Windows shortcuts.

.DESCRIPTION
It can be used to create executable files from PowerShell scripts that do not require a CLI, e.g., background processing scripts and GUI scripts using Microsoft .NET Framework. 

.PARAMETER ScriptFile
TBD

.INPUTS
TBD

.EXAMPLE
TBD

.NOTES
Author: Thales Pinto
Version: 1.2.1
Licence: This code is licensed under the MIT license.
#>

[CmdletBinding(DefaultParameterSetName = "Default shortcut")]

param (
    [Parameter(
        Mandatory = $true,
        ParameterSetName = "PowerShell script",
        HelpMessage = "Enter the path (relative or absolute) of the .ps1 file that you want to create a runnable shortcut from."
    )]
    [validatescript({
        if (-Not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        if ((Get-Item $_).Extension -ne ".ps1") { throw "The source file must end with .ps1." }
        $true
    })]
    [Alias("Script")]
    [string]$ScriptFile,

    [Parameter(
        ParameterSetName = "PowerShell script",
        HelpMessage = "Flag to indicate to pass -NoProfile argument to the PowerShell that will run the script."
    )]
    [switch]$NoProfile,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "Default shortcut",
        HelpMessage = "Enter the path (relative or absolute) to the file for which you want to create a shortcut (.lnk)."
    )]
    [validatescript({
        if (-Not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        $true
    })]
    [string]$SourceFile,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "URL shortcut",
        HelpMessage = "Enter the path (relative or absolute) to the URL shortcut file that you want to create a standard shortcut (.lnk) from."
    )]
    [validatescript({
        if (-Not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        if ((Get-Item $_).Extension -ne ".url") { throw "The source file must end with .url." }
        $true
    })]
    [string]$UrlShortcutFile,

    [Parameter(
        Mandatory = $false,
        ParameterSetName = "Default shortcut"
    )]
    [string]$Arguments,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Enter the path (relative or absolute) to the icon file (.ico) that you want your shortcut to use."
    )]
    [validatescript({
        if (-Not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        if (@(".url", ".ico") -NotContains (Get-Item $_).Extension) { throw "Incompatible file type for icon." }
        $true
    })]
    [Alias("Icon")]
    [string]$IconPath = ",0",

    [Parameter(Mandatory = $false)]
    [string]$Description,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullorEmpty()]
    [Alias("Name", "FinalName")]
    [string]$ShortcutName = "New-Shortcut",

    [Parameter(Mandatory = $false)]
    [validatescript({
        if (-Not (Test-Path -Path $_ -PathType Container)) { throw "Invalid path." }
        $true
    })]
    [string]$OutputPath = $pwd
)

function Get-OutputPath {
    $FinalOutputPath = $OutputPath

    if (-Not ([System.IO.Path]::IsPathRooted($OutputPath))) {
        $FinalOutputPath = $(Resolve-Path -Path $OutputPath)
    }

    $FinalOutputPath = Join-Path $OutputPath -ChildPath "$ShortcutName"
    if (-Not $ShortcutName.EndsWith(".lnk")) {
        $FinalOutputPath += ".lnk"
    }

    return $FinalOutputPath
}

function Get-IconPath {
    if ($IconPath.EndsWith(".lnk")) {
        return $(New-Object -comObject WScript.Shell).CreateShortcut($IconPath).IconLocation
    }
    if ($IconPath.EndsWith(".ico")) {
        return $(Resolve-Path -Path $IconPath).Path
    }

    $FinalIconPath = ",0"

    switch ($PSCmdlet.ParameterSetName) {
        "PowerShell script" {
            if (Test-Path -Path "$PSHome\pwsh.exe" -PathType Leaf) {
                $FinalIconPath = "$PSHome\pwsh.exe"
            } else {
                $FinalIconPath = "$env:systemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
            }
        }
        "URL shortcut" {
            if ($(Get-Content $UrlShortcutFile -Raw) -match "(?im)^\s*IconFile\s*=\s*(.*)") {
                $FinalIconPath = "$($Matches[1] -replace '^\s+|\s+$', ''),0"
            }
        }
    }

    return $FinalIconPath
}

function Test-OutputPathAvailability {
    param (
        [string]$Path
    )

    if (-Not (Test-Path -Path $Path -PathType Leaf)) {
        return $True
    }

    Write-Warning "Shortcut already exists."
    $Question = "Do you want to replace it?"
    $Choices = @(
        [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Replace the old shortcut with the new one")
        [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No changes will be made")
    )

    $Decision = $Host.UI.PromptForChoice("", $Question, $Choices, 1)

    if ($Decision -eq 0) {
        return $True
    }

    return $False
}

$FinalOutputPath = Get-OutputPath
if (-Not (Test-OutputPathAvailability($FinalOutputPath))) {
    Write-Warning "Shortcut not created."
    exit
}

switch ($PSCmdlet.ParameterSetName) {
    "Default shortcut" {
        $TargetPath = $(Resolve-Path -Path $SourceFile)
        $ShortcutArguments = $Arguments
        break
    }
    "PowerShell script" {
        $TargetPath = "$env:ComSpec"
        if ($PSBoundParameters.ContainsKey("NoProfile")) {
            $ShortcutArguments = "/c start /min `"`" pwsh -WindowStyle Hidden -NoProfile -File `"$(Resolve-Path -Path $ScriptFile)`""
        } else {
            $ShortcutArguments = "/c start /min `"`" pwsh -WindowStyle Hidden -File `"$(Resolve-Path -Path $ScriptFile)`""
        }
        break
    }
    "URL shortcut" {
        $TargetPath = "$env:SystemRoot\explorer.exe"
        $ShortcutArguments = (New-Object -ComObject WScript.Shell).CreateShortcut($UrlShortcutFile).TargetPath
        break
    }
}

$Shortcut = $(New-Object -comObject WScript.Shell).CreateShortcut($FinalOutputPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.Arguments = $ShortcutArguments
$Shortcut.IconLocation = Get-IconPath
$Shortcut.Description = $Description
$Shortcut.Save()

if (-Not (Test-Path $Shortcut.FullName)){
    Write-Error "Failed to create shortcut. Try again."
    exit
}

Write-Verbose "Shortcut created with success!"
