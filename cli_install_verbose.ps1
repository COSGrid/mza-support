# Install_MZA.ps1
# Installs COSGrid MicroZAccess MSI with verbose logging enabled for troubleshooting

$msiPath = "C:\Users\vigne\Downloads\COSGrid_MicroZAccess_Windows_V2.1.15.msi"
$logPath = "C:\Users\vigne\Downloads\msi_install.log"

Write-Host "=== Installing COSGrid MicroZAccess ===" -ForegroundColor Cyan

if (-not (Test-Path $msiPath)) {
    Write-Host "MSI file not found at: $msiPath" -ForegroundColor Red
    exit 1
}

# Unblock file (in case it's downloaded from web)
Unblock-File -Path $msiPath

# Run the installer elevated
Start-Process "msiexec.exe" -ArgumentList "/i `"$msiPath`" /L*v `"$logPath`" ALLUSERS=1" -Verb RunAs

Write-Host "`nInstallation started. Check the log file at: $logPath" -ForegroundColor Green
 