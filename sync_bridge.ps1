# ================================================================
# sync_bridge.ps1 — INTAX 대시보드 로컬 동기화 브리지
#
# 포트 19423에서 HTTP 요청을 수신하여 대시보드의 동기화 버튼과
# hometax_auto_sync.ps1 스크립트를 연결합니다.
# INTAX_대시보드_열기.bat에 의해 백그라운드로 자동 시작됩니다.
#
# 엔드포인트:
#   GET /ping    — 서버 동작 확인
#   GET /sync    — 홈택스 동기화 시작 (백그라운드 실행)
#   GET /status  — 동기화 진행 상태 조회
#   GET /reset   — 상태 초기화 (오류 복구용)
# ================================================================

$Port        = 19423
$Root        = $PSScriptRoot
$SyncScript  = Join-Path $Root "hometax_auto_sync.ps1"
$StatusFile  = Join-Path $Root ".sync_status.json"

# ── 이미 실행 중이면 종료 ──────────────────────────────────────
function Test-BridgeRunning {
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:$Port/ping" `
            -UseBasicParsing -TimeoutSec 2 -ErrorAction SilentlyContinue
        return ($r.StatusCode -eq 200)
    } catch { return $false }
}

if (Test-BridgeRunning) {
    Write-Host "[INTAX Bridge] 이미 포트 $Port 에서 실행 중입니다. 종료합니다."
    exit 0
}

# ── 상태 파일 초기화 ──────────────────────────────────────────
@{ status="idle"; message="대기 중"; updatedAt="" } |
    ConvertTo-Json -Compress |
    Set-Content -Path $StatusFile -Encoding UTF8 -Force

# ── HTTP 리스너 시작 ──────────────────────────────────────────
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
try {
    $listener.Start()
} catch {
    Write-Host "[INTAX Bridge] 포트 $Port 시작 실패: $_"
    exit 1
}

Write-Host "[INTAX Bridge] 포트 $Port 에서 대기 중..."

# ── 응답 헬퍼 ────────────────────────────────────────────────
function Send-Json {
    param($ctx, [hashtable]$obj, [int]$code = 200)
    $json  = $obj | ConvertTo-Json -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $ctx.Response.StatusCode  = $code
    $ctx.Response.ContentType = "application/json; charset=utf-8"
    $ctx.Response.Headers.Add("Access-Control-Allow-Origin",  "*")
    $ctx.Response.Headers.Add("Access-Control-Allow-Methods", "GET, OPTIONS")
    $ctx.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
    try   { $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) }
    catch { }
    finally { try { $ctx.Response.Close() } catch { } }
}

function Write-Status {
    param([string]$status, [string]$message)
    @{ status=$status; message=$message; updatedAt=(Get-Date -Format "o") } |
        ConvertTo-Json -Compress |
        Set-Content -Path $StatusFile -Encoding UTF8 -Force
}

function Remove-DoneJobs {
    Get-Job -ErrorAction SilentlyContinue |
        Where-Object { $_.State -in @("Completed","Failed","Stopped") } |
        Remove-Job -Force -ErrorAction SilentlyContinue
}

# ── 메인 루프 ─────────────────────────────────────────────────
try {
    while ($listener.IsListening) {

        $ctx    = $listener.GetContext()
        $path   = $ctx.Request.Url.AbsolutePath
        $method = $ctx.Request.HttpMethod

        # CORS preflight
        if ($method -eq "OPTIONS") {
            Send-Json $ctx @{ ok=$true }
            continue
        }

        switch ($path) {

            "/ping" {
                Send-Json $ctx @{ ok=$true; service="INTAX-Bridge"; port=$Port }
            }

            "/sync" {
                $st = Get-Content $StatusFile -Raw | ConvertFrom-Json
                if ($st.status -eq "running") {
                    Send-Json $ctx @{ status="running"; message="동기화가 이미 진행 중입니다." }
                    break
                }
                if (-not (Test-Path $SyncScript)) {
                    Send-Json $ctx @{ status="error"; message="hometax_auto_sync.ps1을 찾을 수 없습니다." } 500
                    break
                }

                Write-Status "running" "홈택스 동기화를 시작합니다 (2~4분 소요)..."

                $sf = $StatusFile
                $ss = $SyncScript
                Start-Job -ScriptBlock {
                    param($script, $sf)
                    try {
                        $out      = & powershell.exe -ExecutionPolicy Bypass -NonInteractive -File $script 2>&1
                        $exitCode = $LASTEXITCODE
                        if ($exitCode -eq 0) {
                            @{ status="done"; message="동기화 완료"; exitCode=0; updatedAt=(Get-Date -Format "o") } |
                                ConvertTo-Json -Compress | Set-Content $sf -Encoding UTF8 -Force
                        } else {
                            $errLine = ($out | Where-Object { $_ -match "\[ERR\]" } | Select-Object -Last 1) -replace "^\s+\[ERR\]\s*",""
                            $msg = if ($errLine) { $errLine } else { "동기화 실패 (코드: $exitCode)" }
                            @{ status="error"; message=$msg; exitCode=$exitCode; updatedAt=(Get-Date -Format "o") } |
                                ConvertTo-Json -Compress | Set-Content $sf -Encoding UTF8 -Force
                        }
                    } catch {
                        @{ status="error"; message="예외: $_"; exitCode=-1; updatedAt=(Get-Date -Format "o") } |
                            ConvertTo-Json -Compress | Set-Content $sf -Encoding UTF8 -Force
                    }
                } -ArgumentList $ss, $sf | Out-Null

                Remove-DoneJobs
                Send-Json $ctx @{ status="started"; message="홈택스 동기화가 시작됐습니다. 2~4분 소요됩니다." }
            }

            "/status" {
                try {
                    $st = Get-Content $StatusFile -Raw | ConvertFrom-Json
                    Send-Json $ctx @{ status=$st.status; message=$st.message; updatedAt=$st.updatedAt }
                } catch {
                    Send-Json $ctx @{ status="idle"; message="상태 파일 읽기 실패" }
                }
                Remove-DoneJobs
            }

            "/reset" {
                Write-Status "idle" "대기 중"
                Send-Json $ctx @{ ok=$true }
            }

            "/verify-status" {
                # ── 국세청 사업자 상태 조회 API ──────────────────────
                # .env 및 .env(hometax).txt 순서로 NTS_API_KEY 탐색
                $ntsKey = $null
                @(".env", ".env(hometax).txt") | ForEach-Object {
                    if (-not $ntsKey) {
                        $f = Join-Path $Root $_
                        if (Test-Path $f) {
                            (Get-Content $f -Encoding UTF8) | ForEach-Object {
                                if ($_ -match "^NTS_API_KEY=(.+)") { $ntsKey = $Matches[1].Trim() }
                            }
                        }
                    }
                }
                if (-not $ntsKey) {
                    Send-Json $ctx @{ status="error"; message=".env 또는 .env(hometax).txt 파일에 NTS_API_KEY가 없습니다." } 400
                    break
                }

                # POST body 읽기 (JSON 배열: ["1234567890", ...])
                $bodyText = ""
                try {
                    $sr = New-Object System.IO.StreamReader($ctx.Request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyText = $sr.ReadToEnd(); $sr.Close()
                } catch {}

                $bNosRaw = @()
                try { $bNosRaw = @($bodyText | ConvertFrom-Json) } catch {}
                $bNos = @($bNosRaw | ForEach-Object { ($_ -replace "[^0-9]","") } |
                    Where-Object { $_.Length -eq 10 } | Select-Object -Unique)

                if ($bNos.Count -eq 0) {
                    Send-Json $ctx @{ status="error"; message="유효한 사업자번호가 없습니다." } 400
                    break
                }

                $resultMap = @{}
                $errMsg    = $null
                $i = 0
                while ($i -lt $bNos.Count -and -not $errMsg) {
                    $end   = [Math]::Min($i + 99, $bNos.Count - 1)
                    $batch = @($bNos[$i..$end])
                    $reqBody = (@{ b_no = $batch } | ConvertTo-Json -Compress)
                    try {
                        $apiResp = Invoke-RestMethod `
                            -Uri "https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=$ntsKey" `
                            -Method POST -Body $reqBody `
                            -ContentType "application/json; charset=utf-8" `
                            -TimeoutSec 30
                        if ($apiResp.data) {
                            $apiResp.data | ForEach-Object {
                                $st = switch ($_.b_stt_cd) {
                                    "01" { "정상" }
                                    "02" { "휴업" }
                                    "03" { "폐업" }
                                    default { $null }
                                }
                                if ($st) { $resultMap[$_.b_no] = $st }
                            }
                        }
                    } catch {
                        $errMsg = "NTS API 오류: $($_.Exception.Message)"
                    }
                    $i += 100
                }

                if ($errMsg) {
                    Send-Json $ctx @{ status="error"; message=$errMsg } 500
                } else {
                    Send-Json $ctx @{ status="ok"; data=$resultMap; verified=$resultMap.Count }
                }
            }

            default {
                Send-Json $ctx @{ error="알 수 없는 경로: $path" } 404
            }
        }
    }
} finally {
    try { $listener.Stop(); $listener.Close() } catch { }
    Write-Host "[INTAX Bridge] 서버 종료됨."
}
