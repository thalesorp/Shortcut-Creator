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
Version: 1.1.1
Licence: This code is licensed under the MIT license.
#>

[CmdletBinding(DefaultParameterSetName = "Default Shortcut")]

param (
    [Parameter(
        Mandatory = $true,
        ParameterSetName = "PowerShell script",
        HelpMessage = "Enter the path (relative or absolute) of the .ps1 file that you want to create a runnable shortcut from."
    )]
    [validatescript({
        if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
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
        ParameterSetName = "Default Shortcut",
        HelpMessage = "Enter the path (relative or absolute) to the file for which you want to create a shortcut (.lnk)."
    )]
    [validatescript({
        if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        $true
    })]
    [string]$SourceFile,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "URL shortcut",
        HelpMessage = "Enter the path (relative or absolute) to the URL shortcut file that you want to create a standard shortcut (.lnk) from."
    )]
    [validatescript({
        if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        if ((Get-Item $_).Extension -ne ".url") { throw "The source file must end with .url." }
        $true
    })]
    [string]$UrlShortcutFile,

    [Parameter(
        Mandatory = $false,
        ParameterSetName = "Default Shortcut"
    )]
    [string]$Arguments,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Enter the path (relative or absolute) to the icon file (.ico) that you want your shortcut to use."
    )]
    [validatescript({
        if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) { throw "File does not exist." }
        if ((Get-Item $_).Extension -ne ".ico") { throw "The file must be an .ico icon." }
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
        if (-not (Test-Path -Path $_ -PathType Container)) { throw "Invalid path." }
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

    if (Test-Path -Path $FinalOutputPath -PathType Leaf) {
        Write-Warning "Shortcut already exists."
        $Question = "Do you want to replace it?"
        $Choices = @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Replace the old shortcut with the new one")
            [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No changes will be made")
        )

        $Decision = $Host.UI.PromptForChoice("", $Question, $Choices, 1)

        if ($Decision -eq 1) {
            exit
        }
    }

    return $FinalOutputPath
}

function Get-IconPath {
    if ($IconPath -ne ",0") {
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

$FinalOutputPath = Get-OutputPath
$FinalIconPath = Get-IconPath

$Shortcut = $(New-Object -comObject WScript.Shell).CreateShortcut($FinalOutputPath)

switch ($PSCmdlet.ParameterSetName) {
    "Default Shortcut" {
        $Shortcut.TargetPath = $(Resolve-Path -Path $SourceFile)
        $Shortcut.Arguments = $Arguments
        break
    }
    "PowerShell script" {
        $Shortcut.TargetPath = "$env:ComSpec"
        if ($PSBoundParameters.ContainsKey("NoProfile")) {
            $Shortcut.Arguments = "/c start /min `"`" pwsh -WindowStyle Hidden -NoProfile -File `"$(Resolve-Path -Path $ScriptFile)`""
        } else {
            $Shortcut.Arguments = "/c start /min `"`" pwsh -WindowStyle Hidden -File `"$(Resolve-Path -Path $ScriptFile)`""
        }
        break
    }
    "URL shortcut" {
        $Shortcut.TargetPath = "$env:SystemRoot\explorer.exe"
        $Shortcut.Arguments = (New-Object -ComObject WScript.Shell).CreateShortcut($UrlShortcutFile).TargetPath
        break
    }
}

$Shortcut.IconLocation = $FinalIconPath
$Shortcut.Description = $Description
$Shortcut.Save()

if (-not (Test-Path $Shortcut.FullName)){
    Write-Error "Failed to create shortcut."
    exit
}

Write-Verbose "Shortcut created with success!"
