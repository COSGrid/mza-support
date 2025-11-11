# CheckDefenderStatus.ps1
# Displays status of Controlled Folder Access and Attack Surface Reduction rules
Write-Host "=== Checking Microsoft Defender Status ===" -ForegroundColor Cyan

$mpPref = Get-MpPreference | Select EnableControlledFolderAccess, AttackSurfaceReductionRules_Actions

if ($mpPref.EnableControlledFolderAccess -eq 0) {
    Write-Host "Controlled Folder Access: Disabled" -ForegroundColor Yellow
} elseif ($mpPref.EnableControlledFolderAccess -eq 1) {
    Write-Host "Controlled Folder Access: Enabled" -ForegroundColor Green
} else {
    Write-Host "Controlled Folder Access: Unknown ($($mpPref.EnableControlledFolderAccess))" -ForegroundColor Red
}

if ($mpPref.AttackSurfaceReductionRules_Actions) {
    Write-Host "`nAttack Surface Reduction (ASR) rules actions:" -ForegroundColor Cyan
    $mpPref.AttackSurfaceReductionRules_Actions
} else {
    Write-Host "No ASR rules configured." -ForegroundColor Green
}
