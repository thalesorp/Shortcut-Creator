# Shortcut Creator
# Script in PowerShell to create Windows shortcuts. It can be used to create executable files from PowerShell scripts that do not require a CLI, e.g., background processing scripts and GUI scripts using Microsoft .NET Framework. 
#
# 1.0.1
#
# Copyright (c) 2023 Thales Pinto
# This code is licensed under the MIT license.
# See the file LICENSE in the project root for full license information.
#

[CmdletBinding(DefaultParameterSetName = "Default Shortcut")]

param (
    [Parameter(
            Mandatory = $true,
            ParameterSetName = "PowerShell script",
            HelpMessage = "Enter the path (relative or absolute) of the .ps1 file that you want to create a runnable shortcut from."
        )]
        [validatescript({
            if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) {
                throw "File does not exist."
            }
            if ((Get-Item $_).Extension -ne ".ps1") {
                throw "The source file must end with .ps1."
            }
            $true
        })]
        [Alias("Script")]
        [string]
    $ScriptFile,

    [Parameter(
            Mandatory = $true,
            ParameterSetName = "Default Shortcut",
            HelpMessage = "Enter the path (relative or absolute) to the file for which you want to create a shortcut (.lnk)."
        )]
        [validatescript({
            if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) {
                throw "File does not exist."
            }
            $true
        })]
        [string]
    $SourceFile,

    [Parameter(
            Mandatory = $true,
            ParameterSetName = "URL shortcut",
            HelpMessage = "Enter the path (relative or absolute) to the URL shortcut file that you want to create a standard shortcut (.lnk) from."
        )]
        [validatescript({
            if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) {
                throw "File does not exist."
            }
            if ((Get-Item $_).Extension -ne ".url") {
                throw "The source file must end with .url."
            }
            $true
        })]
        [string]
    $UrlShortcutFile,

    [Parameter(
            Mandatory = $false,
            ParameterSetName = "Default Shortcut"
        )]
        [string]
    $Arguments,

    [Parameter(
            Mandatory = $false,
            HelpMessage = "Enter the path (relative or absolute) to the icon file (.ico) that you want your shortcut to use."
        )]
        [validatescript({
            if (-not (Test-Path -Path (Get-Item $_) -PathType Leaf)) {
                throw "File does not exist."
            }
            if ((Get-Item $_).Extension -ne ".ico") {
                throw "The file must be an .ico icon."
            }
            $true
        })]
        [Alias("Icon")]
        [string]
    $IconPath = ",0",

    [Parameter(
            Mandatory = $false
        )]
        [string]
    $Description,

    [Parameter(
            Mandatory = $false
        )]
        [ValidateNotNullorEmpty()]
        [Alias("Name", "FinalName")]
        [string]
    $ShortcutName = "New-Shortcut",

    [Parameter(
            Mandatory = $false
        )]
        [validatescript({
            if (-not (Test-Path -Path $_ -PathType Container)) {
                throw "Invalid path."
            }
            $true
        })]
        [string]
    $OutputPath = $pwd
)


function Test-OutputPath {
    Param (
        [Parameter(Mandatory = $true)][string]$Path
    )

    if (Test-Path -Path $Path -PathType Leaf) {
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
}

if ($PSBoundParameters.ContainsKey("OutputPath")) {
    $OutputPath = $(Resolve-Path -Path $OutputPath)
}

if ($PSBoundParameters.ContainsKey("IconPath")) {
    $IconPath = $(Resolve-Path -Path $IconPath).Path
}

if (-Not $ShortcutName.EndsWith(".lnk")) {
    $FinalOutputPath = "$OutputPath\$ShortcutName.lnk"
} else {
    $FinalOutputPath = "$OutputPath\$ShortcutName"
}

Test-OutputPath($FinalOutputPath)

$Shortcut = $(New-Object -comObject WScript.Shell).CreateShortcut($FinalOutputPath)

switch ($PSCmdlet.ParameterSetName) {
    "Default Shortcut" {
        $Shortcut.TargetPath = $(Resolve-Path -Path $SourceFile)
        $Shortcut.Arguments = $Arguments
        break
    }
    "PowerShell script" {
        $Shortcut.TargetPath = "$env:ComSpec"
        $Shortcut.Arguments = "/c start /min `"`" pwsh -WindowStyle Hidden -File `"$(Resolve-Path -Path $ScriptFile)`""
        break
    }
    "URL shortcut" {
        $Shortcut.TargetPath = "$env:SystemRoot\explorer.exe"
        $Shortcut.Arguments = (New-Object -ComObject WScript.Shell).CreateShortcut($UrlShortcutFile).TargetPath
        break
    }
}

$Shortcut.IconLocation = $IconPath
$Shortcut.Description = $Description
$Shortcut.Save()

if (-not (Test-Path $Shortcut.FullName)){
    Write-Error -Message "Failed to create shortcut."
    exit
}

Write-Output "Shortcut created with success!"
