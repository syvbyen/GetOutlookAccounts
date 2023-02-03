Import-Module (Resolve-Path("$PSSCriptRoot/Modules/GetOutlookAccounts.psm1")) -Force

$getOutlookAccounts = New-GetOutlookAccounts

$getOutlookAccounts.run();