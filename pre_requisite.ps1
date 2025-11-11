#Requires -Version 5.1

<#
.SYNOPSIS
    Pre-installation checker for COSGrid ZTNA MicroZAccess MSI
.DESCRIPTION
    Validates system prerequisites before installing COSGrid ZTNA agent
    Generates detailed audit logs in JSON and text formats
#>

param(
    [string]$OutputPath = "$env:USERPROFILE\Desktop\COSGrid_PreInstall_Audit",
    [string]$ProductName = "COSGrid",
    [int64]$RequiredDiskSpaceMB = 500,
    [string]$InstallDirectory = "C:\Program Files (x86)\COSGrid"
)

# Create output directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = "${OutputPath}_${timestamp}"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$jsonOutputFile = Join-Path $outputDir "audit_report.json"
$textOutputFile = Join-Path $outputDir "audit_report.txt"

# Initialize results object
$auditResults = @{
    AuditTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    ComputerName = $env:COMPUTERNAME
    UserName = $env:USERNAME
    ProductName = $ProductName
    OverallStatus = "PENDING"
    Checks = @{}
}

# Helper function to add check result
function Add-CheckResult {
    param(
        [string]$CheckName,
        [string]$Status,
        [string]$Message,
        [hashtable]$Details = @{}
    )
    
    $auditResults.Checks[$CheckName] = @{
        Status = $Status
        Message = $Message
        Details = $Details
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
    
    Write-Host "[$Status] $CheckName - $Message" -ForegroundColor $(
        switch($Status) {
            "PASS" { "Green" }
            "FAIL" { "Red" }
            "WARNING" { "Yellow" }
            default { "White" }
        }
    )
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "COSGrid ZTNA Pre-Installation Audit" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# 1. CHECK ADMIN RIGHTS
Write-Host "[1/11] Checking Administrator Rights..." -ForegroundColor Cyan
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if ($isAdmin) {
        Add-CheckResult -CheckName "AdminRights" -Status "PASS" -Message "Running with Administrator privileges" -Details @{
            IsAdmin = $true
            CurrentUser = $env:USERNAME
            Domain = $env:USERDOMAIN
        }
    } else {
        Add-CheckResult -CheckName "AdminRights" -Status "FAIL" -Message "Administrator privileges required. Please run as Administrator" -Details @{
            IsAdmin = $false
            CurrentUser = $env:USERNAME
            Domain = $env:USERDOMAIN
        }
    }
} catch {
    Add-CheckResult -CheckName "AdminRights" -Status "FAIL" -Message "Error checking admin rights: $($_.Exception.Message)"
}

# 2. CHECK DISK SPACE
Write-Host "[2/11] Checking Disk Space..." -ForegroundColor Cyan
try {
    $installDrive = $InstallDirectory.Substring(0,2)
    $disk = Get-PSDrive -Name $installDrive.Replace(":", "") -ErrorAction Stop
    $freeSpaceGB = [math]::Round($disk.Free / 1GB, 2)
    $requiredSpaceGB = [math]::Round($RequiredDiskSpaceMB / 1024, 2)
    
    if ($disk.Free -gt ($RequiredDiskSpaceMB * 1MB)) {
        Add-CheckResult -CheckName "DiskSpace" -Status "PASS" -Message "Sufficient disk space available" -Details @{
            Drive = $installDrive
            FreeSpaceGB = $freeSpaceGB
            RequiredSpaceGB = $requiredSpaceGB
            TotalSpaceGB = [math]::Round($disk.Used / 1GB + $freeSpaceGB, 2)
        }
    } else {
        Add-CheckResult -CheckName "DiskSpace" -Status "FAIL" -Message "Insufficient disk space" -Details @{
            Drive = $installDrive
            FreeSpaceGB = $freeSpaceGB
            RequiredSpaceGB = $requiredSpaceGB
            ShortfallGB = [math]::Round($requiredSpaceGB - $freeSpaceGB, 2)
        }
    }
} catch {
    Add-CheckResult -CheckName "DiskSpace" -Status "FAIL" -Message "Error checking disk space: $($_.Exception.Message)"
}

# 3. CHECK FIREWALL RULES
Write-Host "[3/11] Checking Firewall Configuration..." -ForegroundColor Cyan
try {
    $firewallProfiles = Get-NetFirewallProfile -ErrorAction Stop
    $firewallDetails = @{}
    
    foreach ($profile in $firewallProfiles) {
        $firewallDetails[$profile.Name] = @{
            Enabled = $profile.Enabled
            DefaultInboundAction = $profile.DefaultInboundAction
            DefaultOutboundAction = $profile.DefaultOutboundAction
        }
    }
    
    $existingRules = Get-NetFirewallRule -DisplayName "*COSGrid*", "*ZTNA*", "*MicroZAccess*" -ErrorAction SilentlyContinue
    
    Add-CheckResult -CheckName "FirewallRules" -Status "PASS" -Message "Firewall configuration retrieved successfully" -Details @{
        Profiles = $firewallDetails
        ExistingCOSGridRules = ($existingRules | Measure-Object).Count
        Note = "Installation may need to create firewall rules"
    }
} catch {
    Add-CheckResult -CheckName "FirewallRules" -Status "WARNING" -Message "Could not fully check firewall: $($_.Exception.Message)"
}

# 4. CHECK SYSTEM ARCHITECTURE
Write-Host "[4/11] Checking System Architecture..." -ForegroundColor Cyan
try {
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $csInfo = Get-CimInstance -ClassName Win32_ComputerSystem
    $arch = $env:PROCESSOR_ARCHITECTURE
    
    $details = @{
        OSArchitecture = $osInfo.OSArchitecture
        ProcessorArchitecture = $arch
        SystemType = $csInfo.SystemType
        OSVersion = $osInfo.Version
        OSName = $osInfo.Caption
        BuildNumber = $osInfo.BuildNumber
    }
    
    if ($arch -eq "AMD64" -or $arch -eq "x64") {
        Add-CheckResult -CheckName "SystemArchitecture" -Status "PASS" -Message "64-bit system detected" -Details $details
    } elseif ($arch -eq "x86") {
        Add-CheckResult -CheckName "SystemArchitecture" -Status "WARNING" -Message "32-bit system detected - verify MSI compatibility" -Details $details
    } else {
        Add-CheckResult -CheckName "SystemArchitecture" -Status "WARNING" -Message "Unusual architecture detected: $arch" -Details $details
    }
} catch {
    Add-CheckResult -CheckName "SystemArchitecture" -Status "FAIL" -Message "Error checking system architecture: $($_.Exception.Message)"
}

# 5. CHECK PREVIOUS VERSION
Write-Host "[5/11] Checking for Previous Versions..." -ForegroundColor Cyan
try {
    $installedProducts = @()
    
    # Check 64-bit registry
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($path in $regPaths) {
        $products = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -like "*COSGrid*" -or $_.DisplayName -like "*ZTNA*" -or $_.DisplayName -like "*MicroZAccess*" }
        
        foreach ($product in $products) {
            $installedProducts += @{
                Name = $product.DisplayName
                Version = $product.DisplayVersion
                Publisher = $product.Publisher
                InstallDate = $product.InstallDate
                UninstallString = $product.UninstallString
            }
        }
    }
    
    if ($installedProducts.Count -eq 0) {
        Add-CheckResult -CheckName "PreviousVersion" -Status "PASS" -Message "No previous versions detected" -Details @{
            ProductsFound = 0
        }
    } else {
        Add-CheckResult -CheckName "PreviousVersion" -Status "WARNING" -Message "Previous version(s) found - may need uninstall first" -Details @{
            ProductsFound = $installedProducts.Count
            Products = $installedProducts
        }
    }
} catch {
    Add-CheckResult -CheckName "PreviousVersion" -Status "WARNING" -Message "Error checking previous versions: $($_.Exception.Message)"
}

# 6. CHECK CONFLICTING SOFTWARE
Write-Host "[6/11] Checking for Conflicting Software..." -ForegroundColor Cyan
try {
    $conflictingKeywords = @("VPN", "Zero Trust", "ZTNA", "Remote Access")
    $conflictingSoftware = @()
    
    foreach ($path in $regPaths) {
        foreach ($keyword in $conflictingKeywords) {
            $products = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName -like "*$keyword*" -and $_.DisplayName -notlike "*COSGrid*" }
            
            foreach ($product in $products) {
                $conflictingSoftware += @{
                    Name = $product.DisplayName
                    Version = $product.DisplayVersion
                    Publisher = $product.Publisher
                    Category = $keyword
                }
            }
        }
    }
    
    # Remove duplicates
    $conflictingSoftware = $conflictingSoftware | Select-Object -Property Name, Version, Publisher, Category -Unique
    
    if ($conflictingSoftware.Count -eq 0) {
        Add-CheckResult -CheckName "ConflictingSoftware" -Status "PASS" -Message "No obvious conflicting software detected" -Details @{
            ConflictsFound = 0
        }
    } else {
        Add-CheckResult -CheckName "ConflictingSoftware" -Status "WARNING" -Message "Potentially conflicting software detected" -Details @{
            ConflictsFound = $conflictingSoftware.Count
            Software = $conflictingSoftware
            Note = "Review if these applications may conflict with ZTNA agent"
        }
    }
} catch {
    Add-CheckResult -CheckName "ConflictingSoftware" -Status "WARNING" -Message "Error checking conflicting software: $($_.Exception.Message)"
}

# 7. CHECK WINDOWS INSTALLER SERVICE
Write-Host "[7/11] Checking Windows Installer Service..." -ForegroundColor Cyan
try {
    $msiService = Get-Service -Name "msiserver" -ErrorAction Stop
    
    $details = @{
        ServiceName = $msiService.Name
        DisplayName = $msiService.DisplayName
        Status = $msiService.Status
        StartType = $msiService.StartType
        CanStop = $msiService.CanStop
    }
    
    if ($msiService.Status -eq "Running" -or $msiService.StartType -ne "Disabled") {
        Add-CheckResult -CheckName "WindowsInstallerService" -Status "PASS" -Message "Windows Installer service is available" -Details $details
    } else {
        Add-CheckResult -CheckName "WindowsInstallerService" -Status "FAIL" -Message "Windows Installer service is disabled" -Details $details
    }
} catch {
    Add-CheckResult -CheckName "WindowsInstallerService" -Status "FAIL" -Message "Error checking Windows Installer service: $($_.Exception.Message)"
}

# 8. CHECK REGISTRY ACCESS PERMISSIONS
Write-Host "[8/11] Checking Registry Access Permissions..." -ForegroundColor Cyan
try {
    $testPaths = @(
        "HKLM:\SOFTWARE",
        "HKLM:\SYSTEM\CurrentControlSet\Services"
    )
    
    $accessResults = @{}
    $allAccessible = $true
    
    foreach ($path in $testPaths) {
        try {
            $acl = Get-Acl -Path $path -ErrorAction Stop
            $accessResults[$path] = @{
                Accessible = $true
                Owner = $acl.Owner
            }
        } catch {
            $accessResults[$path] = @{
                Accessible = $false
                Error = $_.Exception.Message
            }
            $allAccessible = $false
        }
    }
    
    if ($allAccessible) {
        Add-CheckResult -CheckName "RegistryAccessPermissions" -Status "PASS" -Message "Registry access permissions verified" -Details $accessResults
    } else {
        Add-CheckResult -CheckName "RegistryAccessPermissions" -Status "FAIL" -Message "Limited registry access detected" -Details $accessResults
    }
} catch {
    Add-CheckResult -CheckName "RegistryAccessPermissions" -Status "FAIL" -Message "Error checking registry permissions: $($_.Exception.Message)"
}

# 9. CHECK FILE SYSTEM PERMISSIONS
Write-Host "[9/11] Checking File System Permissions..." -ForegroundColor Cyan
try {
    $testPaths = @(
        "C:\Program Files",
        "C:\Windows\System32",
        $env:TEMP
    )
    
    $fsResults = @{}
    $allWritable = $true
    
    foreach ($path in $testPaths) {
        $testFile = Join-Path $path "cosgrid_test_$(Get-Random).tmp"
        try {
            [System.IO.File]::WriteAllText($testFile, "test")
            Remove-Item $testFile -Force -ErrorAction SilentlyContinue
            $fsResults[$path] = @{
                Writable = $true
            }
        } catch {
            $fsResults[$path] = @{
                Writable = $false
                Error = $_.Exception.Message
            }
            if ($path -ne $env:TEMP) {
                $allWritable = $false
            }
        }
    }
    
    if ($allWritable) {
        Add-CheckResult -CheckName "FileSystemPermissions" -Status "PASS" -Message "File system permissions verified" -Details $fsResults
    } else {
        Add-CheckResult -CheckName "FileSystemPermissions" -Status "FAIL" -Message "Limited file system access detected" -Details $fsResults
    }
} catch {
    Add-CheckResult -CheckName "FileSystemPermissions" -Status "FAIL" -Message "Error checking file system permissions: $($_.Exception.Message)"
}

# 10. CHECK PENDING RESTARTS
Write-Host "[10/11] Checking for Pending Restarts..." -ForegroundColor Cyan
try {
    $pendingReboot = $false
    $rebootReasons = @()
    
    # Check Component Based Servicing
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
        $pendingReboot = $true
        $rebootReasons += "Component Based Servicing"
    }
    
    # Check Windows Update
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
        $pendingReboot = $true
        $rebootReasons += "Windows Update"
    }
    
    # Check PendingFileRenameOperations
    $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    if ($pfro) {
        $pendingReboot = $true
        $rebootReasons += "Pending File Rename Operations"
    }
    
    # Check System Center Configuration Manager
    try {
        $ccmReboot = Invoke-WmiMethod -Namespace "root\ccm\ClientSDK" -Class "CCM_ClientUtilities" -Name "DetermineIfRebootPending" -ErrorAction SilentlyContinue
        if ($ccmReboot.RebootPending) {
            $pendingReboot = $true
            $rebootReasons += "SCCM Client"
        }
    } catch {
        # SCCM not installed, ignore
    }
    
    $details = @{
        PendingReboot = $pendingReboot
        Reasons = $rebootReasons
        LastBootTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    }
    
    if (-not $pendingReboot) {
        Add-CheckResult -CheckName "PendingRestarts" -Status "PASS" -Message "No pending restarts detected" -Details $details
    } else {
        Add-CheckResult -CheckName "PendingRestarts" -Status "WARNING" -Message "System restart pending" -Details $details
    }
} catch {
    Add-CheckResult -CheckName "PendingRestarts" -Status "WARNING" -Message "Error checking pending restarts: $($_.Exception.Message)"
}

# 11. CHECK ACL RULES / GROUP POLICIES
Write-Host "[11/11] Checking ACL Rules / Group Policies..." -ForegroundColor Cyan
try {
    $gpResult = @{}
    
    # Check if system is domain-joined
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $isDomainJoined = $computerSystem.PartOfDomain
    
    $gpResult["DomainJoined"] = $isDomainJoined
    $gpResult["Domain"] = $computerSystem.Domain
    
    # Check software restriction policies
    $srpPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers"
    if (Test-Path $srpPath) {
        $gpResult["SoftwareRestrictionPolicies"] = "Configured"
    } else {
        $gpResult["SoftwareRestrictionPolicies"] = "Not Configured"
    }
    
    # Check AppLocker
    $appLockerService = Get-Service -Name "AppIDSvc" -ErrorAction SilentlyContinue
    if ($appLockerService) {
        $gpResult["AppLocker"] = @{
            ServiceStatus = $appLockerService.Status
            ServiceStartType = $appLockerService.StartType
        }
    } else {
        $gpResult["AppLocker"] = "Not Configured"
    }
    
    # Check user rights assignments
    $userRights = @{
        SeServiceLogonRight = "Log on as a service"
        SeInteractiveLogonRight = "Log on locally"
    }
    
    Add-CheckResult -CheckName "ACLRulesGroupPolicies" -Status "PASS" -Message "Group Policy configuration retrieved" -Details $gpResult
} catch {
    Add-CheckResult -CheckName "ACLRulesGroupPolicies" -Status "WARNING" -Message "Error checking group policies: $($_.Exception.Message)"
}

# DETERMINE OVERALL STATUS
Write-Host "`n========================================" -ForegroundColor Cyan
$failCount = ($auditResults.Checks.Values | Where-Object { $_.Status -eq "FAIL" }).Count
$warningCount = ($auditResults.Checks.Values | Where-Object { $_.Status -eq "WARNING" }).Count
$passCount = ($auditResults.Checks.Values | Where-Object { $_.Status -eq "PASS" }).Count

if ($failCount -eq 0 -and $warningCount -eq 0) {
    $auditResults.OverallStatus = "READY"
    Write-Host "Overall Status: READY TO INSTALL" -ForegroundColor Green
} elseif ($failCount -eq 0) {
    $auditResults.OverallStatus = "READY_WITH_WARNINGS"
    Write-Host "Overall Status: READY WITH WARNINGS" -ForegroundColor Yellow
} else {
    $auditResults.OverallStatus = "NOT_READY"
    Write-Host "Overall Status: NOT READY - CRITICAL ISSUES FOUND" -ForegroundColor Red
}

$auditResults.Summary = @{
    TotalChecks = $auditResults.Checks.Count
    Passed = $passCount
    Warnings = $warningCount
    Failed = $failCount
}

Write-Host "========================================`n" -ForegroundColor Cyan

# SAVE JSON REPORT
Write-Host "Saving JSON report to: $jsonOutputFile" -ForegroundColor Cyan
$auditResults | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonOutputFile -Encoding UTF8

# GENERATE TEXT REPORT
Write-Host "Generating text report to: $textOutputFile" -ForegroundColor Cyan

$textReport = @"
================================================================================
                COSGrid ZTNA MicroZAccess - Pre-Installation Audit Report
================================================================================

Audit Information:
------------------
Timestamp:      $($auditResults.AuditTimestamp)
Computer Name:  $($auditResults.ComputerName)
User Name:      $($auditResults.UserName)
Product:        $($auditResults.ProductName)

Overall Status: $($auditResults.OverallStatus)
------------------
Total Checks:   $($auditResults.Summary.TotalChecks)
Passed:         $($auditResults.Summary.Passed)
Warnings:       $($auditResults.Summary.Warnings)
Failed:         $($auditResults.Summary.Failed)

================================================================================
                              DETAILED CHECK RESULTS
================================================================================

"@

foreach ($checkName in $auditResults.Checks.Keys | Sort-Object) {
    $check = $auditResults.Checks[$checkName]
    $statusSymbol = switch($check.Status) {
        "PASS" { "[PASS]" }
        "FAIL" { "[FAIL]" }
        "WARNING" { "[WARN]" }
        default { "[INFO]" }
    }
    
    $textReport += "`n$statusSymbol $checkName`n"
    $textReport += "Status:  $($check.Status)`n"
    $textReport += "Message: $($check.Message)`n"
    $textReport += "Time:    $($check.Timestamp)`n`n"
    
    if ($check.Details.Count -gt 0) {
        $textReport += "Details:`n"
        foreach ($key in $check.Details.Keys) {
            $value = $check.Details[$key]
            if ($value -is [hashtable] -or $value -is [array]) {
                $textReport += "  $key : $(ConvertTo-Json $value -Compress -Depth 3)`n"
            } else {
                $textReport += "  $key : $value`n"
            }
        }
        $textReport += "`n"
    }
    
    $textReport += "--------------------------------------------------------------------------------`n"
}

$textReport += "`n"
$textReport += "================================================================================"
$textReport += "`n                                 RECOMMENDATIONS"
$textReport += "`n================================================================================"
$textReport += "`n`n"

if ($failCount -gt 0) {
    $textReport += "CRITICAL ISSUES (Must be resolved before installation):`n"
    $textReport += "--------------------------------------------------------`n"
    foreach ($checkName in $auditResults.Checks.Keys | Sort-Object) {
        $check = $auditResults.Checks[$checkName]
        if ($check.Status -eq "FAIL") {
            $textReport += "- $checkName : $($check.Message)`n"
        }
    }
    $textReport += "`n"
}

if ($warningCount -gt 0) {
    $textReport += "WARNINGS (Review before installation):`n"
    $textReport += "--------------------------------------`n"
    foreach ($checkName in $auditResults.Checks.Keys | Sort-Object) {
        $check = $auditResults.Checks[$checkName]
        if ($check.Status -eq "WARNING") {
            $textReport += "- $checkName : $($check.Message)`n"
        }
    }
    $textReport += "`n"
}

if ($failCount -eq 0 -and $warningCount -eq 0) {
    $textReport += "All prerequisite checks passed successfully!`n"
    $textReport += "System is ready for COSGrid ZTNA MicroZAccess installation.`n`n"
}

$textReport += "================================================================================"
$textReport += "`n                                END OF REPORT"
$textReport += "`n================================================================================"
$textReport += "`n`n"
$textReport += "Generated by: COSGrid ZTNA Pre-Installation Checker`n"
$textReport += "Report Location: $outputDir`n"

$textReport | Out-File -FilePath $textOutputFile -Encoding UTF8

# DISPLAY FINAL MESSAGE
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Audit Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Reports saved to: $outputDir" -ForegroundColor White
Write-Host "  - JSON Report: audit_report.json" -ForegroundColor White
Write-Host "  - Text Report: audit_report.txt" -ForegroundColor White
Write-Host "`nOverall Status: $($auditResults.OverallStatus)" -ForegroundColor $(
    switch($auditResults.OverallStatus) {
        "READY" { "Green" }
        "READY_WITH_WARNINGS" { "Yellow" }
        "NOT_READY" { "Red" }
        default { "White" }
    }
)

if ($failCount -gt 0) {
    Write-Host "`nCRITICAL: $failCount check(s) failed. Please resolve these issues before installing." -ForegroundColor Red
}

if ($warningCount -gt 0) {
    Write-Host "`nWARNING: $warningCount check(s) returned warnings. Review before proceeding." -ForegroundColor Yellow
}

Write-Host "`nPress any key to open the report directory..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Start-Process explorer.exe -ArgumentList $outputDir