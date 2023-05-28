# Shortcut Creator

Script in PowerShell to create Windows shortcuts. It can be used to create executable files from PowerShell scripts that do not require a CLI, e.g., background processing scripts and GUI scripts using Microsoft .NET Framework. 

## Usage

### Parameters

- **ScriptFile**: Specifies the path (relative or absolute) of the .ps1 file that you want to create a runnable shortcut from.
- **SourceFile**: Specifies the path (relative or absolute) to the file for which you want to create a shortcut (.lnk).
- **UrlShortcutFile**: Specifies the path (relative or absolute) to the URL shortcut file that you want to create a standard shortcut (.lnk) from.
- **Arguments**: Specifies arguments to use when starting the target of the shortcut.
- **IconPath**: Specifies the path (relative or absolute) to the icon file (.ico or .lnk) that you want your shortcut to use.
- **Description**: Specifies a description for the shortcut.
- **ShortcutName**: Specifies the name of the shortcut.
- **OutputPath**: Specifies the path (relative or absolute) where you want to create the shortcut.

### Examples

Create a shortcut from a PowerShell script:

```powershell
New-Shortcut -ScriptFile "C:\Scripts\MyScript.ps1" -ShortcutName "MyScript"
```

Create a shortcut from an executable file:

```powershell
New-Shortcut -SourceFile "C:\Program Files\MyApp\MyApp.exe" -ShortcutName "MyApp" -IconPath "C:\Program Files\MyApp\MyAppIcon.ico"
```

Create a shortcut from a URL shortcut file on the current directory:

```powershell
New-Shortcut -UrlShortcutFile "C:\Users\Me\Links\MyLink.url" -ShortcutName "MyLink" -OutputPath .
```

## License

This code is licensed under the MIT license. See the file LICENSE in the project root for full license information.
