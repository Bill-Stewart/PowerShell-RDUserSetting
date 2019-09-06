# RDUserSetting.psm1
# Written by Bill Stewart (bstewart@iname.com)
#
# As of this writing, the PowerShell LocalAccount module cannot query or set a
# number of settings related to Remote Desktop Services (RDS) features. These
# settings are found on the following tabs in the "Local Users and Groups"
# console:
#
# * Environment
# * Remote Desktop Services Profile
# * Remote control
# * Sessions
#
# This module rectifies this limitation and lets you query and/or set these
# settings from PowerShell.
#
# The following table lists the available properties and their locations in the
# "Local Users and Groups" console:
#
# Property                          Data Type  GUI Tab                          Setting
# --------                          ---------  -------                          -------
# RDInitialProgramRequired          Boolean    Environment                      "Start the following program at logon"
# RDInitialProgram                  String     Environment                      "Program file name"
# RDInitialProgramWorkingDirectory  String     Environment                      "Start in"
# RDconnectDrives                   Boolean    Environment                      "Connect client drives at logon"
# RDconnectPrinters                 Boolean    Environment                      "Connect client printers at logon"
# RDSetDefaultPrinter               Boolean    Environment                      "Default to main client printer"
# RDProfilePath                     String     Remote Desktop Services Profile  "Profile Path"
# RDHomeDrive                       String     Remote Desktop Services Profile  "Connect"
# RDHomeDirectory                   String     Remote Desktop Services Profile  "Local path" or "To:"
# RDAllowLogon                      Boolean    Remote Desktop Services Profile  "Deny this user permissions to log on to Remote Desktop Session Host server"
# RDRemoteControlSetting            Enum       Remote control                   (See below)
# RDDisconnectedSessionLimit        Integer    Sessions                         "End a disconnected session" (milliseconds)
# RDActiveSessionLimit              Integer    Sessions                         "Active session limit" (milliseconds)
# RDIdleSessionLimit                Integer    Sessions                         "Idle session limit" (milliseconds)
# RDEndSessionIfDisconnected        Boolean    Sessions                         "Disconnect from session" or "End session"
# RDReconnectToNewSession           Boolean    Sessions                         "From any client" or "From originating client only"
#
# The RDRemoteControlSetting property is an Enum that can be any of the
# following values:
#
# Value                     GUI Settings
# -----                     ------------
# * Disabled                "Enable remote control" unchecked
# * NotifyInteractive       "Enable remote control" checked; "Require user's permission" checked; "Interact with the session" selected
# * NoNotifyInteractive     "Enable remote control" checked; "Require user's permission" unchecked; "Interact with the session" selected
# * NotifyNonInteractive    "Enable remote control" checked; "Require user's permission" checked; "View the user's session" selected
# * NoNotifyNonInteractive  "Enable remote control" checked; "Require user's permission" unchecked; "View the user's session" selected
#
# This module exports two functions:
#
# Get-RDUserSetting - outputs the above properties for users
# Set-RDUserSetting - sets one or more of the above properties for users
#
# A few notes about performance:
#
# * Get-RDUserSetting has to call the WTSQueryUserConfig API once for each
#   Remote Desktop property (API limitation). Be aware of this when using this
#   function to get RDS properties for users over slow network links.
#
# * Set-RDUserSetting calls the WTSSetUserConfig API once for each property you
#   set (again, API limitation), so the same caveat applies.
#
# * Set-RDUserSetting with the -PassThru parameter produces the same effect as
#   running Set-RDUserSetting followed by Get-RDUserSetting, so again: be aware
#   that the number of API calls required may not perform well over slow links.
#
# Version history:
# 1.0.0 (2019-09-04)
# * Initial version.

#requires -version 2

# Struct:
# * [DBA9F1BBB58F4527A346681DC006E551.NetApi32+NET_DISPLAY_USER]
# Methods:
# * [DBA9F1BBB58F4527A346681DC006E551.NetApi32]::NetApiBufferFree()
# * [DBA9F1BBB58F4527A346681DC006E551.NetApi32]::NetQueryDisplayInformation()
Add-Type -MemberDefinition @"
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct NET_DISPLAY_USER {
  public string usri1_name;
  public string usri1_comment;
  public uint   usri1_flags;
  public string usri1_full_name;
  public uint   usri1_user_id;
  public uint   usri1_next_index;
}
[DllImport("netapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern uint NetApiBufferFree(IntPtr Buffer);
[DllImport("netapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern uint NetQueryDisplayInformation(
  string ServerNname,
  uint Level,
  uint Index,
  uint EntriesRequested,
  uint PreferredMaximumLength,
  ref uint ReturnedEntryCount,
  ref IntPtr SortedBuffer);
"@ -Namespace DBA9F1BBB58F4527A346681DC006E551 -Name NetApi32

# Enum:
# * [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]
# Methods:
# * [DBA9F1BBB58F4527A346681DC006E551.WtsApi32]::WTSFreeMemory
# * [DBA9F1BBB58F4527A346681DC006E551.WtsApi32]::WTSQueryUserConfig
# * [DBA9F1BBB58F4527A346681DC006E551.WtsApi32]::WtsSetUserConfig
Add-Type -MemberDefinition @"
public enum WTS_CONFIG_CLASS {
  WTSUserConfigInitialProgram,
  WTSUserConfigWorkingDirectory,
  WTSUserConfigfInheritInitialProgram,
  WTSUserConfigfAllowLogonTerminalServer,
  WTSUserConfigTimeoutSettingsConnections,
  WTSUserConfigTimeoutSettingsDisconnections,
  WTSUserConfigTimeoutSettingsIdle,
  WTSUserConfigfDeviceClientDrives,
  WTSUserConfigfDeviceClientPrinters,
  WTSUserConfigfDeviceClientDefaultPrinter,
  WTSUserConfigBrokenTimeoutSettings,
  WTSUserConfigReconnectSettings,
  WTSUserConfigModemCallbackSettings,
  WTSUserConfigModemCallbackPhoneNumber,
  WTSUserConfigShadowingSettings,
  WTSUserConfigTerminalServerProfilePath,
  WTSUserConfigTerminalServerHomeDir,
  WTSUserConfigTerminalServerHomeDirDrive,
  WTSUserConfigfTerminalServerRemoteHomeDir
}
[DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern void WTSFreeMemory(IntPtr pMemory);
[DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool WTSQueryUserConfig(
  string pServerName,
  string pUserName,
  int WTSConfigClass,
  out IntPtr ppBuffer,
  out uint pBytesReturned);
[DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool WTSSetUserConfig(
  string pServerName,
  string pUserName,
  int WTSConfigClass,
  IntPtr pBuffer,
  uint DataLength);
"@ -Namespace DBA9F1BBB58F4527A346681DC006E551 -Name WtsApi32

# Enum:
# * [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting]::Disabled
# * [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting]::NotifyInteractive
# * [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting]::NoNotifyInteractive
# * [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting]::NotifyNonInteractive
# * [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting]::NoNotifyNonInteractive
Add-Type -TypeDefinition @"
namespace DBA9F1BBB58F4527A346681DC006E551 {
  public enum RemoteControlSetting {
    Disabled,
    NotifyInteractive,
    NoNotifyInteractive,
    NotifyNonInteractive,
    NoNotifyNonInteractive
  }
}
"@

# Win32 API values
$ERROR_MORE_DATA      = 234
$MAX_PREFERRED_LENGTH = [BitConverter]::ToUInt32([BitConverter]::GetBytes(-1), 0)
$UF_NORMAL_ACCOUNT    = 512

# Managed return type for each item in the WTS_CONFIG_CLASS enum (used by
# WTSQueryUserConfig and WTSSetUserConfig)
$ConfigClassTypes = @{
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigInitialProgram                = "String"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigWorkingDirectory              = "String"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfInheritInitialProgram        = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfAllowLogonTerminalServer     = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsConnections    = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsDisconnections = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsIdle           = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDrives           = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientPrinters         = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDefaultPrinter   = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigBrokenTimeoutSettings         = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigReconnectSettings             = "UInt32"
# [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigModemCallbackSettings         = "UInt32"  # Not implemented
# [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigModemCallbackPhoneNumber      = "String"  # Not implemented
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigShadowingSettings             = "UInt32"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerProfilePath     = "String"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDir         = "string"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDirDrive    = "String"
  [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfTerminalServerRemoteHomeDir  = "UInt32"
}

# Outputs a custom object based on hashtables
function Out-Object {
  param(
    [Collections.Hashtable[]] $hashData
  )
  $order = @()
  $result = @{}
  foreach ( $item in $hashData ) {
    $order += ($item.Keys -as [Array])[0]
    $result += $item
  }
  New-Object PSObject -Property $result | Select-Object $order
}

# Writes a custom error to the error stream
function Write-CustomError {
  param(
    [Exception] $exception,

    $targetObject,

    [String] $errorID,

    [Management.Automation.ErrorCategory] $errorCategory = "NotSpecified",

    [Switch] $terminatingError
  )
  $errorRecord = New-Object Management.Automation.ErrorRecord $exception,$errorID,$errorCategory,$targetObject
  if ( -not $terminatingError ) {
    $PSCmdlet.WriteError($errorRecord)
  }
  else {
    $PSCmdlet.ThrowTerminatingError($errorRecord)
  }
}

# Executes NetQueryDisplayInformation using P/Invoke.
function NetQueryDisplayInformation {
  param(
    [String] $computerName
  )
  $index = 0
  $returnedEntryCount = 0
  $pSortedBuffer = [IntPtr]::Zero
  do {
    try {
      # Use NetQueryDisplayInformation to retrieve 100 names at a time
      $apiResult = [DBA9F1BBB58F4527A346681DC006E551.NetApi32]::NetQueryDisplayInformation(
        $computerName,              # ServerName
        1,                          # Level
        $index,                     # Index
        100,                        # EntriesRequested
        $MAX_PREFERRED_LENGTH,      # PreferredMaximumLength
        [Ref] $returnedEntryCount,  # ReturnedEntryCount
        [Ref] $pSortedBuffer)       # SortedBuffer
      if ( ($apiResult -eq 0) -or ($apiResult -eq $ERROR_MORE_DATA) ) {
        # Get address of initial entry in buffer
        $offset = $pSortedBuffer.ToInt64()
        for ( ; $returnedEntryCount -gt 0; $returnedEntryCount-- ) {
          # Point at entry
          $pEntry = New-Object IntPtr($offset)
          # Copy unmanaged to managed type
          $netDisplayUser = [Runtime.InteropServices.Marshal]::PtrToStructure($pEntry, [Type] [DBA9F1BBB58F4527A346681DC006E551.NetApi32+NET_DISPLAY_USER])
          # If normal user account, output it
          if ( ($netDisplayUser.usri1_flags -band $UF_NORMAL_ACCOUNT) -ne 0 ) {
            $netDisplayUser.usri1_name
          }
          # Set index for next search
          $index = $netDisplayUser.usri1_next_index
          # Increment offset for next member
          $offset += [Runtime.InteropServices.Marshal]::SizeOf($netDisplayUser)
        }
      }
      else {
        $exception = New-Object ComponentModel.Win32Exception ([Runtime.InteropServices.Marshal]::GetLastWin32Error())
        $errorID = (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name
        Write-CustomError $exception "Computer '$computerName'" $errorID
      }
    }
    catch {
      # Hopefully don't see this (unhandled exception)
      Write-Error -Exception $_.Exception
    }
    finally {
      # Free unmanaged buffer
      [Void] [DBA9F1BBB58F4527A346681DC006E551.NetApi32]::NetApiBufferFree($pSortedBuffer)
    }
  }
  while ( $apiResult -eq $ERROR_MORE_DATA )
}

# Executes WTSQueryUserConfig using P/Invoke.
# Outputs nothing ($null) on failure.
function WTSQueryUserConfig {
  param(
    [String] $computerName,

    [String] $userName,

    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS] $configClass
  )
  try {
    $pBuffer = [IntPtr]::Zero
    $bytesReturned = 0
    $apiSuccess = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32]::WTSQueryUserConfig(
      $computerName,        # pServerName
      $userName,            # pUserName
      $configClass,         # WTSConfigClass
      [Ref] $pBuffer,       # ppBuffer
      [Ref] $bytesReturned  # pBytesReturned
    )
    if ( $apiSuccess ) {
      switch ( $ConfigClassTypes[$configClass] ) {
        "String" {
          # Copy unmanaged to managed string
          [Runtime.InteropServices.Marshal]::PtrToStringAuto($pBuffer)
        }
        "UInt32" {
          # Copy unmanaged to managed int
          [Runtime.InteropServices.Marshal]::ReadInt32($pBuffer) -as [UInt32]
        }
      }
    }
  }
  catch {
    # Hopefully don't see this (unhandled exception)
    Write-Error -Exception $_.Exception
  }
  finally {
    # Free buffer
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32]::WTSFreeMemory($pBuffer)
  }
  if ( -not $apiSuccess ) {
    $exception = New-Object ComponentModel.Win32Exception ([Runtime.InteropServices.Marshal]::GetLastWin32Error())
    $errorID = (Get-Variable MyInvocation -Scope 2).Value.MyCommand.Name
    Write-CustomError $exception "User '$userName' on computer '$computerName'" $errorID
  }
}

# Executes WTSQueryUserConfig API for each WTS type and outputs an object (we
# have to call WTSQueryUserConfig multiple times per user)
function GetRDUserSetting {
  param(
    [String] $computerName,

    [String] $userName
  )
  # First: Find out if home directory uses drive mapping; if we fail, no point
  # in trying to call the API again, so throw an error and end the function
  $homeDirUsesMappedDrive = WTSQueryUserConfig $computerName $userName ([DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfTerminalServerRemoteHomeDir)
  if ( $null -eq $homeDirUsesMappedDrive ) {
    return
  }
  # Call API using these classes (in this order)
  $configClasses = @(
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfInheritInitialProgram
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigInitialProgram
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigWorkingDirectory
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDrives
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientPrinters
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDefaultPrinter
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerProfilePath
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDir
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDirDrive
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfAllowLogonTerminalServer
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigShadowingSettings
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsDisconnections
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsConnections
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsIdle
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigBrokenTimeoutSettings
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigReconnectSettings
  )
  $friendlyNames = @{
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfInheritInitialProgram        = "RDInitialProgramRequired"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigInitialProgram                = "RDInitialProgram"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigWorkingDirectory              = "RDInitialProgramWorkingDirectory"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDrives           = "RDConnectDrives"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientPrinters         = "RDConnectPrinters"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDefaultPrinter   = "RDSetDefaultPrinter"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerProfilePath     = "RDProfilePath"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDirDrive    = "RDHomeDrive"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDir         = "RDHomeDirectory"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfAllowLogonTerminalServer     = "RDAllowLogon"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigShadowingSettings             = "RDRemoteControlSetting"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsDisconnections = "RDDisconnectedSessionLimit"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsConnections    = "RDActiveSessionLimit"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsIdle           = "RDIdleSessionLimit"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigBrokenTimeoutSettings         = "RDEndSessionIfDisconnected"
    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigReconnectSettings             = "RDReconnectToNewSession"
  }
  # Create output object with initial properties
  $outputObject = Out-Object `
    @{"ComputerName"                     = $computerName},  # String
    @{"UserName"                         = $userName},      # String
    @{"RDInitialProgramRequired"         = $null},          # Boolean
    @{"RDInitialProgram"                 = $null},          # String
    @{"RDInitialProgramWorkingDirectory" = $null},          # String
    @{"RDconnectDrives"                  = $null},          # Boolean
    @{"RDconnectPrinters"                = $null},          # Boolean
    @{"RDSetDefaultPrinter"              = $null},          # Boolean
    @{"RDProfilePath"                    = $null},          # String
    @{"RDHomeDrive"                      = $null},          # String
    @{"RDHomeDirectory"                  = $null},          # String
    @{"RDAllowLogon"                     = $null},          # Boolean
    @{"RDRemoteControlSetting"           = $null},          # DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting
    @{"RDDisconnectedSessionLimit"       = $null},          # UInt32
    @{"RDActiveSessionLimit"             = $null},          # UInt32
    @{"RDIdleSessionLimit"               = $null},          # Uint32
    @{"RDEndSessionIfDisconnected"       = $null},          # Boolean
    @{"RDReconnectToNewSession"          = $null}           # Boolean
  # Set the properties in the output object
  foreach ( $configClass in $configClasses ) {
    $result = WTSQueryUserConfig $computerName $userName $configClass
    # Stop trying if we fail
    if ( $null -eq $result ) {
      break
    }
    switch -Regex ( $friendlyNames[$configClass] ) {
      '^RDInitialProgramRequired$' {
        # Inverse of other int values (0 = required, 1 = not required)
        $outputObject.$_ = -not ($result -as [Boolean])
        break
      }
      '(^RDInitialProgram$)|(^RDInitialProgramWorkingDirectory$)' {
        # Only meaningful if initial program is required
        if ( $outputObject.RDInitialProgramRequired ) {
          $outputObject.$_ = $result
        }
        break
      }
      '(^RDConnectDrives$)|(^RDConnectPrinters$)|(^RDSetDefaultPrinter$)|(^RDAllowLogon$)|(^RDEndSessionIfDisconnected$)|(^RDReconnectToNewSession$)' {
        # Boolean
        $outputObject.$_ = $result -as [Boolean]
        break
      }
      '(^RDProfilePath$)|(^RDHomeDirectory$)' {
        # String
        $outputObject.$_ = $result
        break
      }
      '^RDHomeDrive$' {
        # Only meaningful if drive mapping in use
        if ( $homeDirUsesMappedDrive -as [Boolean] ) {
          $outputObject.$_ = $result
        }
        break
      }
      '^RDRemoteControlSetting$' {
        # Enum
        $outputObject.$_ = $result -as [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting]
        break
      }
      '(^RDDisconnectedSessionLimit$)|(^RDActiveSessionLimit$)|(^RDIdleSessionLimit$)' {
        # Int
        $outputObject.$_ = $result
        break
      }
    }
  }
  $outputObject
}

# Executes WTSSetUserConfig API using P/Invoke.
# Outputs 0 for success, non-zero for failure.
function WTSSetUserConfig {
  param(
    [String] $computerName,

    [String] $userName,

    [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS] $configClass,

    $newValue
  )
  switch ( $ConfigClassTypes[$configClass] ) {
    "String" {
      # Allocate unmanaged memory and copy string
      $pNewValue = [Runtime.InteropServices.Marshal]::StringToHGlobalAuto([String] $newValue)
      if ( $newValue.Length -gt 0 ) {
        $dataLength = $newValue.Length
      }
      else {
        # 1 for empty string (single null character)
        $dataLength = 1
      }
    }
    "UInt32" {
      # Allocate unmanaged memory
      $pNewValue = [Runtime.InteropServices.Marshal]::AllocHGlobal([Runtime.InteropServices.Marshal]::SizeOf([Type] [UInt32]))
      # Copy managed variable to unmanaged memory
      [Runtime.InteropServices.Marshal]::WriteInt32($pNewValue, [UInt32] $newValue)
      # Data length is size of data referenced by pointer
      $dataLength = [Runtime.InteropServices.Marshal]::SizeOf($pNewValue)
    }
  }
  try {
    $apiSuccess = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32]::WTSSetUserConfig(
      $computerName,  # pServerName
      $userName,      # pUserName
      $configClass,   # WTSConfigClass
      $pNewValue,     # pBuffer
      $dataLength     # DataLength
    )
  }
  catch {
    # Hopefully don't see this (unhandled exception)
    Write-Error -Exception $_.Exception
  }
  finally {
    # Free unmanaged memory
    [Runtime.InteropServices.Marshal]::FreeHGlobal($pNewValue)
  }
  if ( $apiSuccess ) {
    return 0
  }
  else {
    $exception = New-Object ComponentModel.Win32Exception ([Runtime.InteropServices.Marshal]::GetLastWin32Error())
    $errorID = (Get-Variable MyInvocation -Scope 2).Value.MyCommand.Name
    Write-CustomError $exception "User '$userName' on computer '$computerName'" $errorID
    return $exception.ErrorCode
  }
}

# Executes WTSSetUserConfig API for each property specified in the hashtable.
function SetRDUserSetting {
  param(
    [String] $computerName,

    [String] $userName,

    [Collections.Hashtable] $parameters
  )
  $friendlyNames = @{
    "rdInitialProgramRequired"         = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfInheritInitialProgram
    "rdInitialProgram"                 = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigInitialProgram
    "rdInitialProgramWorkingDirectory" = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigWorkingDirectory
    "rdConnectDrives"                  = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDrives
    "rdConnectPrinters"                = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientPrinters
    "rdSetDefaultPrinter"              = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfDeviceClientDefaultPrinter
    "rdProfilePath"                    = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerProfilePath
    "rdHomeDrive"                      = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDirDrive
    "rdHomeDirectory"                  = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTerminalServerHomeDir
    "rdAllowLogon"                     = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigfAllowLogonTerminalServer
    "rdRemoteControlSetting"           = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigShadowingSettings
    "rdDisconnectedSessionLimit"       = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsDisconnections
    "rdActiveSessionLimit"             = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsConnections
    "rdIdleSessionLimit"               = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigTimeoutSettingsIdle
    "rdEndSessionIfDisconnected"       = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigBrokenTimeoutSettings
    "rdReconnectToNewSession"          = [DBA9F1BBB58F4527A346681DC006E551.WtsApi32+WTS_CONFIG_CLASS]::WTSUserConfigReconnectSettings
  }
  foreach ( $parameterName in $parameters.Keys ) {
    switch -Regex ( $parameterName ) {
      '^rdInitialProgramRequired$' {
        # Inverse of other int values (0 = required, 1 = not required)
        $result = WTSSetUserConfig $computerName $userName $friendlyNames[$_] ((-not $parameters[$parameterName]) -as [UInt32])
        break
      }
      '(^rdInitialProgram$)|(^rdInitialProgramWorkingDirectory$)|(^rdProfilePath$)|(^rdHomeDrive$)|(^rdHomeDirectory$)' {
        # String values
        $result = WTSSetUserConfig $computerName $userName $friendlyNames[$_] $parameters[$parameterName]
        break
      }
      '(^rdConnectDrives$)|(^rdConnectPrinters$)|(^rdSetDefaultPrinter$)|(^RDAllowLogon$)|(^rdRemoteControlSetting$)|(^rdDisconnectedSessionLimit$)|(^rdActiveSessionLimit$)|(^rdIdleSessionLimit$)|(^rdEndSessionIfDisconnected$)|(^rdReconnectToNewSession$)' {
        # Int values
        $result = WTSSetUserConfig $computerName $userName $friendlyNames[$_] ($parameters[$parameterName] -as [UInt32])
        break
      }
    }
    if ( $result -ne 0 ) {
      break
    }
  }
  return $result
}

# Exported function
function Get-RDUserSetting {
  <#
  .SYNOPSIS
  Gets Remote Desktop user settings for one or more users.

  .DESCRIPTION
  Gets Remote Desktop user settings for one or more users.

  .PARAMETER UserName
  Specifies one or more user names. Wildcards are not supported. If you omit this parameter, all users are assumed.

  .PARAMETER ComputerName
  Specifies one or more computer names. The default is the current computer. Wildcards are not supported. An empty string ("") or a single dot (".") refer to the current computer.

  .OUTPUTS
  Objects with the following properties:
    ComputerName - Computer name where user is located
    UserName - User name
    RDInitialProgramRequired - Whether initial program is required
    RDInitialProgram - Name of initial program
    RDInitialProgramWorkingDirectory - Working directory for initial program
    RDconnectDrives - Whether to connect client drives at logon
    RDconnectPrinters - Whether to connect client printers at logon
    RDSetDefaultPrinter - Whether to set default printer at logon
    RDProfilePath - Path of Remote Desktop Services (RDS) user profile
    RDHomeDrive - Mapped RDS home directory drive letter (e.g., "H:")
    RDHomeDirectory - Path of RDS home directory
    RDAllowLogon - Whether user can logon using RDS
    RDRemoteControlSetting - Remote control settings for user*
    RDDisconnectedSessionLimit - End disconnected sessions after (milliseconds)
    RDActiveSessionLimit - Maximum sesssion duration (milliseconds)
    RDIdleSessionLimit - Maximum idle session time (milliseconds)
    RDEndSessionIfDisconnected - End session if disconnected
    RDReconnectToNewSession - Whether to reconnect to previous sessions
  * Possible values are Disabled, NotifyInteractive, NoNotifyInteractive,
  NotifyNonInteractive, or NoNotifyNonInteractive

  .EXAMPLE
  PS C:\> Get-RDUserSetting
  This command outputs the Remote Deskop settings for all users on the current computer.

  .EXAMPLE
  PS C:\> Get-RDUserSetting | Select-Object UserName,RDRemoteControlSetting
  This command outputs the username and Remote Destkop remote control setting for all users on the current computer.

  .EXAMPLE
  PS C:\> "SERVER1","SERVER2" | ForEach-Object { Get-RDUserSetting -ComputerName $_ }
  This command outputs the Remote Desktop settings for all users on the computers SERVER1 and SERVER2.

  .EXAMPLE
  PS C:\> Get-Content UserList.txt | Get-RDUserSetting | Export-Csv RDSUserInfo.txt -NoTypeInformation
  This command gets the Remote Desktop settings for each user listed in the file UserList.txt and outputs the results to a CSV file.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Position = 0,ValueFromPipeline = $true)]
    [String[]] $UserName,

    [Parameter(Position = 1)]
    [String[]] $ComputerName = ""
  )
  process {
    foreach ( $computerNameItem in $ComputerName ) {
      if ( ($computerNameItem -eq "") -or ($computerNameItem -eq ".") ) {
        $computerNameItem = [Net.Dns]::GetHostName()
      }
      if ( ($UserName | Measure-Object).Count -gt 0 ) {
        foreach ( $userNameItem in $UserName ) {
          GetRDUserSetting $computerNameItem $userNameItem
        }
      }
      else {
        # Use ForEach-Object here in case of a large number of users
        NetQueryDisplayInformation $computerNameItem | ForEach-Object {
          GetRDUserSetting $computerNameItem $_
        }
      }
    }
  }
}

# Exported function
function Set-RDUserSetting {
  <#
  .SYNOPSIS
  Sets Remote Desktop user settings for one or more users.

  .DESCRIPTION
  Sets Remote Desktop user settings for one or more users.

  .PARAMETER UserName
  Specifies one or more user names. Wildcards are not supported.

  .PARAMETER ComputerName
  Specifies one or more computer names. The default is the current computer. Wildcards are not supported. An empty string ("") or a single dot (".") refer to the current computer.

  .PARAMETER RDInitialProgramRequired
  Specifies whether an initial program is required when logging onto Remote Desktop Services (RDS).

  .PARAMETER RDInitialProgram
  Specifies the name of the RDS initial program, if an initial program is required. Can be an empty string ("").

  .PARAMETER RDInitialProgramWorkingDirectory
  Specifies the working directory for the RDS initial program, if initial program is required. Can be an empty string ("").

  .PARAMETER RDconnectDrives
  Specifies the user's drives should be reconnected when logging on via RDS.

  .PARAMETER RDconnectPrinters
  Specifies whether the user's printers should be reconnected when logging on via RDS.

  .PARAMETER RDSetDefaultPrinter
  Specifies whether the user's default printer will be set when logging on via RDS.

  .PARAMETER RDProfilePath
  Specifies the path of the user's RDS user profile. Can be an empty string ("").

  .PARAMETER RDHomeDrive
  Specifies the drive letter for the user's RDS home directory. Can be an empty string (""), but if specified, this value must be a drive letter followed by ":" (e.g., "H:").

  .PARAMETER RDHomeDirectory
  Specicies the path of of the user's RDS home directory. Can be an empty string ("").

  .PARAMETER RDAllowLogon
  Specifies whether the user can logon using RDS.

  .PARAMETER RDRemoteControlSetting
  Specifies the RDS remote control settings for the user: Disabled, NotifyInteractive (user is notified and session is interactive), NoNotifyInteractive (user is not notified and session is interactive), NotifyNonInteractive (user is notified and session is not interactive), or NoNotifyNonInteractive (user is not notified and session is not interactive).

  .PARAMETER RDDisconnectedSessionLimit
  Specifies that the user's disconnected RDS sessions should be terminated after this many milliseconds. A value of zero (0) prevents the user's disconnected RDS sessions from being terminated.

  .PARAMETER RDActiveSessionLimit
  Specifies the maximum number of milliseconds the user is allowed to be connected via RDS. A value of zero (0) allows the user to stay connected indefinitely.

  .PARAMETER RDIdleSessionLimit
  Specifies the maximum number of milliseconds the user's RDS session is allowed to remain idle, after which the session will be disconnected (if RDEndSessionIfDisconnected is $false) or terminated (if RDEndSessionIfDisconnected is $true).

  .PARAMETER RDEndSessionIfDisconnected
  Specifies whether the user's idle or broken RDS session will be disconnected ($false) or terminated ($true).

  .PARAMETER RDReconnectToNewSession
  Specifies how a user can reconnected to a previous session. If $false, the user can use any client computer to reconnect to a disconnected session; if $true, the user can only reconnect to a previous session using the same client computer. (If the user logs on from a different client computer and this setting is $true, the user will get a new logon session.)

  .PARAMETER PassThru
  Specifies to output an object representing the user's new settings (by default, there is no output). See the help for Get-RDUserSettings for information about output objects.

  .EXAMPLE
  PS C:\> Set-RDUserSetting KenDyer -RDAllowLogon $false
  This command prohibits the KenDyer account from logging on using Remote Desktop.

  .EXAMPLE
  PS C:\> Get-RDUserSetting | Where-Object { $_.RDRemoteControlSetting -ne "Disabled" } | ForEach-Object { Set-RDUserSetting -RDRemoteControlSettings "Disabled" }
  This command finds all users on the current computer that do not have Remote Desktop remote control disabled, and disables the setting for only those users.

  .EXAMPLE
  PS C:\> "SERVER1","SERVER2" | ForEach-Object { Get-RDUserSetting -ComputerName $_ } | Where-Object { -not $_.RDAllowLogon } | ForEach-Object { Set-RDUserSetting -RDAllowLogon $true }
  This command finds all users on SERVER1 and SERVER2 that are currently disabled for Remote Desktop logon, and enables the setting for only those users.

  .EXAMPLE
  PS C:\> Get-Content UserList.txt | Set-RDUserSetting -RDInitialProgramRequired $true -RDInitialProgram "C:\Program Files\Internet Explorer\iexplore.exe"
  This command sets the initial Remote Desktop program for each of the users in the file UserList.txt on the current computer.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Position = 0,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [String[]] $UserName,

    [Parameter(Position = 1,ValueFromPipelineByPropertyName = $true)]
    [String[]] $ComputerName = "",

    [Boolean] $RDInitialProgramRequired,

    [String] $RDInitialProgram,

    [String] $RDInitialProgramWorkingDirectory,

    [Boolean] $RDConnectDrives,

    [Boolean] $RDConnectPrinters,

    [Boolean] $RDSetDefaultPrinter,

    [String] $RDProfilePath,

    [String] $RDHomeDrive,

    [String] $RDHomeDirectory,

    [Boolean] $RDAllowLogon,

    [DBA9F1BBB58F4527A346681DC006E551.RemoteControlSetting] $RDRemoteControlSetting,

    [UInt32] $RDDisconnectedSessionLimit,

    [UInt32] $RDActiveSessionLimit,

    [UInt32] $RDIdleSessionLimit,

    [Boolean] $RDEndSessionIfDisconnected,

    [Boolean] $RDReconnectToNewSession,

    [Switch] $PassThru
  )
  begin {
    $availableParameters = @(
      "RDInitialProgramRequired","RDInitialProgram",
      "RDInitialProgramWorkingDirectory","RDConnectDrives","RDConnectPrinters",
      "RDSetDefaultPrinter","RDProfilePath","RDHomeDrive","RDHomeDirectory",
      "RDAllowLogon","RDRemoteControlSetting","RDDisconnectedSessionLimit",
      "RDActiveSessionLimit","RDIdleSessionLimit","RDEndSessionIfDisconnected",
      "RDReconnectToNewSession"
    )
    $specifiedParameters = @{}
    foreach ( $parameterName in $PSBoundParameters.Keys ) {
      if ( $availableParameters -contains $parameterName ) {
        $specifiedParameters.Add($parameterName, $PSBoundParameters.Item($parameterName))
      }
    }
  }
  process {
    if ( ($UserName | Measure-Object).Count -gt 0 ) {
      if ( $specifiedParameters.Count -gt 0 ) {
        foreach ( $computerNameItem in $ComputerName ) {
          if ( ($computerNameItem -eq "") -or ($computerNameItem -eq ".") ) {
            $computerNameItem = [Net.Dns]::GetHostName()
          }
          else {
            # Is computer name a property?
            if ( $computerNameItem.ComputerName ) {
              $computerNameItem = $computerNameItem.ComputerName
            }
          }
          foreach ( $userNameItem in $userName ) {
            # Is user name a property?
            if ( $userNameItem.UserName ) {
              $userNameItem = $userNameItem.UserName
            }
            if ( $PSCmdlet.ShouldProcess("User '$userNameItem' on computer '$computerNameItem'", "Set Remote Desktop user settings") ) {
              $result = SetRDUserSetting $computerNameItem $userNameItem $specifiedParameters
              if ( $PassThru -and ($result -eq 0) ) {
                GetRDUserSetting $computerNameItem $userNameItem
              }
            }
          }
        }
      }
      else {
        $OFS = ", "
        $exception = New-Object Management.Automation.ParameterBindingException "Missing one or more parameters. Specify one of the following parameters and try again: $availableParameters"
        Write-CustomError $exception "" $MyInvocation.MyCommand.Name "InvalidArgument" -terminatingError
      }
    }
    else {
      $exception = New-Object Management.Automation.ParameterBindingException "Argument for parameter 'UserName' is null or empty."
      Write-CustomError $exception "" $MyInvocation.MyCommand.Name "InvalidData" -terminatingError
    }
  }
}
