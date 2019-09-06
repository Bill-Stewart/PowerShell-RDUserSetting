@{
  # ModuleToProcess = 'RDUserSetting.psm1'
  RootModule        = 'RDUserSetting.psm1'
  ModuleVersion     = '1.0.0.0'
  GUID              = '36bd6c4e-a68f-4727-8df5-11d2323caa3b'
  Author            = 'Bill Stewart'
  CompanyName       = 'Bill Stewart'
  Copyright         = '(C) 2019 by Bill Stewart'
  Description       = 'Get and set Remote Desktop (RDS) user settings.'
  PowerShellVersion = '3.0'
  AliasesToExport   = '*'
  FunctionsToExport = @(
    'Get-RDUserSetting'
    'Set-RDUserSetting'
  )
}
