@{
  ModuleToProcess   = 'RDUserSetting.psm1'
  # RootModule      = 'RDUserSetting.psm1'
  ModuleVersion     = '1.0.0.0'
  GUID              = '1f7ce1e0-a188-4c15-8fb1-36e8284d4385'
  Author            = 'Bill Stewart'
  CompanyName       = 'Bill Stewart'
  Copyright         = '(C) 2019 by Bill Stewart'
  Description       = 'Get and set Remote Desktop (RDS) user settings.'
  PowerShellVersion = '2.0'
  AliasesToExport   = '*'
  FunctionsToExport = @(
    'Get-RDUserSetting'
    'Set-RDUserSetting'
  )
}
