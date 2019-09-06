#define ModuleName "RDUserSetting"
#define AppName ModuleName + " PowerShell Module"
#define AppPublisher "Bill Stewart"
#define AppVersion "1.0"
#define InstallPath "WindowsPowerShell\Modules\" + ModuleName
#define InstallPathPS2 "WindowsPowerShell\v1.0\Modules\" + ModuleName
#define IconFilename ModuleName + ".ico"
#define SetupCompany "Bill Stewart (bstewart@iname.com)"
#define SetupVersion "1.0.0.0"

[Setup]
AppId={{D4C5077B-D5F1-4ECF-80AE-9E4F0E32FAE5}
AppName={#AppName}
AppPublisher={#AppPublisher}
AppVersion={#AppVersion}
ArchitecturesInstallIn64BitMode=x64
Compression=lzma2/max
DefaultDirName={code:GetInstallDir}
DisableDirPage=yes
MinVersion=6.0
OutputBaseFilename={#ModuleName}_{#AppVersion}_Setup
OutputDir=.
PrivilegesRequired=admin
SetupIconFile={#IconFilename}
SolidCompression=yes
UninstallDisplayIcon={code:GetInstallDir}\{#IconFilename}
UninstallFilesDir={code:GetInstallDir}\Uninstall
VersionInfoCompany={#SetupCompany}
VersionInfoProductVersion={#AppVersion}
VersionInfoVersion={#SetupVersion}
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardResizable=no
WizardSizePercent=150
WizardSmallImageFile={#ModuleName}_55x55.bmp
WizardStyle=modern

[Languages]
Name: english; InfoBeforeFile: "Readme.rtf"; LicenseFile: "License.rtf"; MessagesFile: "compiler:Default.isl"

[Files]
; PSv2, 32-bit
Source: "{#IconFilename}";        DestDir: "{syswow64}\{#InstallPathPS2}";  Check: (not IsWindowsPowerShell3OrLater) and (not Is64BitInstallMode)
Source: "License.txt";            DestDir: "{syswow64}\{#InstallPathPS2}";  Check: not IsWindowsPowerShell3OrLater
Source: "Readme.md";              DestDir: "{syswow64}\{#InstallPathPS2}";  Check: not IsWindowsPowerShell3OrLater
Source: "{#ModuleName}_PS2.psd1"; DestDir: "{syswow64}\{#InstallPathPS2}";  Check: not IsWindowsPowerShell3OrLater; DestName: "{#ModuleName}.psd1"
Source: "{#ModuleName}.psm1";     DestDir: "{syswow64}\{#InstallPathPS2}";  Check: not IsWindowsPowerShell3OrLater
; PSv2, 64-bit
Source: "{#IconFilename}";        DestDir: "{sysnative}\{#InstallPathPS2}"; Check: (not IsWindowsPowerShell3OrLater) and Is64BitInstallMode
Source: "License.txt";            DestDir: "{sysnative}\{#InstallPathPS2}"; Check: (not IsWindowsPowerShell3OrLater) and Is64BitInstallMode
Source: "Readme.md";              DestDir: "{sysnative}\{#InstallPathPS2}"; Check: (not IsWindowsPowerShell3OrLater) and Is64BitInstallMode
Source: "{#ModuleName}_PS2.psd1"; DestDir: "{sysnative}\{#InstallPathPS2}"; Check: (not IsWindowsPowerShell3OrLater) and Is64BitInstallMode; DestName: "{#ModuleName}.psd1"
Source: "{#ModuleName}.psm1";     DestDir: "{sysnative}\{#InstallPathPS2}"; Check: (not IsWindowsPowerShell3OrLater) and Is64BitInstallMode
; PSv3+, 32-bit
Source: "{#IconFilename}";        DestDir: "{commonpf32}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater and (not Is64BitInstallMode)
Source: "License.txt";            DestDir: "{commonpf32}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater
Source: "Readme.md";              DestDir: "{commonpf32}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater
Source: "{#ModuleName}.psd1";     DestDir: "{commonpf32}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater
Source: "{#ModuleName}.psm1";     DestDir: "{commonpf32}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater
; PSv3+, 64-bit
Source: "{#IconFilename}";        DestDir: "{commonpf64}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater and Is64BitInstallMode
Source: "License.txt";            DestDir: "{commonpf64}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater and Is64BitInstallMode
Source: "Readme.md";              DestDir: "{commonpf64}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater and Is64BitInstallMode
Source: "{#ModuleName}.psd1";     DestDir: "{commonpf64}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater and Is64BitInstallMode
Source: "{#ModuleName}.psm1";     DestDir: "{commonpf64}\{#InstallPath}";   Check: IsWindowsPowerShell3OrLater and Is64BitInstallMode

[Code]
Var
  WindowsPowerShellMajorVersion: Integer;

Function GetWindowsPowerShellMajorVersion(): Integer;
  Var
    RootPath, VersionString: String;
    SubkeyNames: TArrayOfString;
    HighestPSVersion, I, PSVersion: Integer;
  Begin
  Result := 0;
  RootPath := 'SOFTWARE\Microsoft\PowerShell';
  If Not RegGetSubkeyNames(HKEY_LOCAL_MACHINE, RootPath, SubkeyNames) Then
    Exit;
  HighestPSVersion := 0;
  For I := 0 To GetArrayLength(SubkeyNames) - 1 Do
    Begin
    If RegQueryStringValue(HKEY_LOCAL_MACHINE, RootPath + '\' + SubkeyNames[I] + '\PowerShellEngine', 'PowerShellVersion', VersionString) Then
      Begin
      PSVersion := StrToIntDef(Copy(VersionString, 0, 1), 0);
      If PSVersion > HighestPSVersion Then
        HighestPSVersion := PSVersion;
      End;
    End;
  Result := HighestPSVersion;
  End;

Function InitializeSetup(): Boolean;
  Begin
  WindowsPowerShellMajorVersion := GetWindowsPowerShellMajorVersion();
  Result := WindowsPowerShellMajorVersion > 1;
  If Not Result Then
    Begin
    Log('FATAL: Setup cannot continue because Windows PowerShell version 2.0 or later is required.');
    If Not WizardSilent() Then
      Begin
      MsgBox('Setup cannot continue because Windows PowerShell version 2.0 or later is required.'
        + #13#10#13#10 + 'Setup will now exit.', mbCriticalError, MB_OK);
      Exit;
      End;
    End;
  Log('Windows PowerShell major version detected: ' + IntToStr(WindowsPowerShellMajorVersion));
  Result := True;
  End;

Function IsWindowsPowerShell3OrLater(): Boolean;
  Begin
  Result := WindowsPowerShellMajorVersion > 2;
  End;

Function GetInstallDir(Param: String): String;
  Begin
  If Not IsWindowsPowerShell3OrLater() Then
    Begin
    If Not Is64BitInstallMode() Then
      Begin
      // PSv2, 32-bit
      Result := ExpandConstant('{syswow64}\{#InstallPathPS2}');
      End
    Else
      Begin
      // PSv2, 64-bit
      Result := ExpandConstant('{sysnative}\{#InstallPathPS2}');
      End;
    End
  Else
    Begin
    If Not Is64BitInstallMode() Then
      Begin
      // PSv3+, 32-bit
      Result := ExpandConstant('{commonpf32}\{#InstallPath}');
      End
    Else
      Begin
      // PSv3+, 64-bit
      Result := ExpandConstant('{commonpf64}\{#InstallPath}');
      End;
    End;
  End;
