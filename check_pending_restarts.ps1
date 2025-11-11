# CheckPendingRename.ps1
# Checks if any files are pending rename (which often means a reboot is required)

Write-Host "=== Checking for Pending File Rename Operations ===" -ForegroundColor Cyan

try {
    $pendingOps = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction Stop
    if ($pendingOps.PendingFileRenameOperations) {
        Write-Host "`nPending operations detected:" -ForegroundColor Yellow
        $pendingOps.PendingFileRenameOperations
    } else {
        Write-Host "`nNo pending file rename operations found." -ForegroundColor Green
    }
}
catch {
    Write-Host "No PendingFileRenameOperations key found - no reboot required." -ForegroundColor Green
}
