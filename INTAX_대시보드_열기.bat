@echo off
chcp 65001 >nul

:: ── 동기화 브리지 서버 백그라운드 시작 (이미 실행 중이면 자동 무시) ──
start /b "" powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File "%~dp0sync_bridge.ps1"
timeout /t 2 /nobreak >nul

:: ── INTAX 대시보드 열기 ──────────────────────────────────────
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="%LOCALAPPDATA%\INTAXDashboard" "%~dp0index.html"
