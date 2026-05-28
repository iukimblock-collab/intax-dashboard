# ================================================================
# setup_autostart.ps1 — INTAX 브리지 서버 Windows 자동 시작 등록
#
# [실행 방법]
#   PowerShell을 관리자 권한으로 열고:
#   .\setup_autostart.ps1
#
# [등록 내용]
#   - 작업 이름: INTAX-Dashboard-Bridge
#   - 트리거: Windows 로그인 시 자동 시작
#   - 실행: sync_bridge.ps1 (숨김 창)
#   - 재시작: 오류 시 3회 자동 재시작
# ================================================================

$Root       = $PSScriptRoot
$BridgeScript = Join-Path $Root "sync_bridge.ps1"
$TaskName   = "INTAX-Dashboard-Bridge"

if (-not (Test-Path $BridgeScript)) {
    Write-Host "[오류] sync_bridge.ps1을 찾을 수 없습니다: $BridgeScript" -ForegroundColor Red
    exit 1
}

Write-Host "INTAX 브리지 서버 자동 시작 등록 중..." -ForegroundColor Cyan

# 기존 작업 제거
try {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$BridgeScript`""

$trigger = New-ScheduledTaskTrigger -AtLogOn

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 12) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -MultipleInstances IgnoreNew

try {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -RunLevel Highest `
        -Force | Out-Null

    Write-Host "[완료] '$TaskName' 작업이 등록됐습니다." -ForegroundColor Green
    Write-Host "  → 다음 Windows 로그인부터 브리지 서버가 자동으로 시작됩니다." -ForegroundColor White
    Write-Host ""
    Write-Host "지금 바로 시작하려면:" -ForegroundColor Yellow
    Write-Host "  Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
    Write-Host "  또는 INTAX_대시보드_열기.bat 실행" -ForegroundColor White

} catch {
    Write-Host "[오류] 작업 등록 실패: $_" -ForegroundColor Red
    Write-Host "PowerShell을 관리자 권한으로 실행하세요." -ForegroundColor Yellow
    exit 1
}
