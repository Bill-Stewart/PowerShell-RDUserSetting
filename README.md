# RDUserSetting PowerShell Module

This PowerShell module allows getting and setting of Remote Desktop (RD) user properties. Commands:

**Get-RDUserSetting  
Set-RDUserSetting**

This module is mainly useful for managing the RD properties of local accounts on non-server versions of Windows (because non-server versions of Windows don't expose the RD properties in the **Local Users and Groups** console).

## System Requirements

* Windows Vista or Windows Server 2008 or later
* Windows PowerShell 2.0 or later

On Windows Vista and Windows Server 2008, Windows Management Framework (WMF) 2.0 is also required to meet the Windows PowerShell prerequisite.

## Properties

The following table lists the properties you can get and set:

| Property                         | GUI Tab                         | GUI Setting
| -------------------------------- | ------------------------------- | -----------
| RDInitialProgramRequired         | Environment                     | **Start the following program at logon**
| RDInitialProgram                 | Environment                     | **Program file name**
| RDInitialProgramWorkingDirectory | Environment                     | **Start in**
| RDconnectDrives                  | Environment                     | **Connect client drives at logon**
| RDconnectPrinters                | Environment                     | **Connect client printers at logon**
| RDSetDefaultPrinter              | Environment                     | **Default to main client printer**
| RDProfilePath                    | Remote Desktop Services Profile | **Profile path**
| RDHomeDrive                      | Remote Desktop Services Profile | **Connect**
| RDHomeDirectory                  | Remote Desktop Services Profile | *(path of home directory or mapped drive)*
| RDAllowLogon                     | Remote Desktop Services Profile | **Deny this user permissions to log on to Remote Desktop Session Host server**
| RDRemoteControlSetting           | Remote control                  | *(See below)*
| RDDisconnectedSessionLimit       | Sessions                        | **End a disconnected session** *(in milliseconds)*
| RDActiveSessionLimit             | Sessions                        | **Active session limit** *(in milliseconds)*
| RDIdleSessionLimit               | Sessions                        | **Idle session limit** *(in milliseconds)*
| RDEndSessionIfDisconnected       | Sessions                        | **Disconnect from session or **End session**
| RDReconnectToNewSession          | Sessions                        | **From any client** or **From originating client only**

The following table shows the possible enumeration values for the **RDRemoteControlSetting** property and the corresponding settings on the **Remote control** GUI tab:

|                            | Enable remote control | Require user's permission | View the user's session | Interact with the session
| -------------------------- | --------------------- | ------------------------- | ----------------------- | -------------------------
| **Disabled**               | Unchecked             | --                        | --                      | --
| **NotifyInteractive**      | Checked               | Checked                   | Not selected            | Selected
| **NoNotifyInteractive**    | Checked               | Unchecked                 | Not selected            | Selected
| **NotifyNonInteractive**   | Checked               | Checked                   | Selected                | Not selected
| **NoNotifyNonInteractive** | Checked               | Unchecked                 | Selected                | Not selected


## Contributions

If you would like to contribute to this project, use this link:

https://paypal.me/wastewart

Thank you!