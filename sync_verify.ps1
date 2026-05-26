# ================================================================
# sync_verify.ps1
# 홈택스 자동 동기화 완료 후 빠른 sanity check
#
# 사용법:
#   .\sync_verify.ps1
#   .\sync_verify.ps1 -PreviousCount 570   # 이전 거래처 수 명시
#   .\sync_verify.ps1 -MaxAgeMinutes 30    # 최신성 기준 분 변경 (기본 10분)
#
# PowerShell 5.1 호환 (&&, ??, ?. 미사용)
# ================================================================

param(
    [int]$PreviousCount  = 0,
    [int]$MaxAgeMinutes  = 10
)

$Root       = $PSScriptRoot
$BackupJson = Join-Path $Root "intax_backup.json"
$RawUrl     = "https://raw.githubusercontent.com/iukimblock-collab/intax-dashboard/main/intax_backup.json"

# ── 출력 헬퍼 ─────────────────────────────────────────────────
$script:PassCount = 0
$script:FailCount = 0
$script:Results   = @()

function Write-Pass([string]$msg) {
    $line = "[PASS] $msg"
    Write-Host $line -ForegroundColor Green
    $script:PassCount++
    $script:Results += $line
}

function Write-Fail([string]$msg) {
    $line = "[FAIL] $msg"
    Write-Host $line -ForegroundColor Red
    $script:FailCount++
    $script:Results += $line
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  sync_verify.ps1 — 동기화 완료 후 빠른 검증" -ForegroundColor Cyan
Write-Host "  실행 시각: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "================================================================"
Write-Host ""

# ── V1. intax_backup.json 수정 시각 < MaxAgeMinutes 이내 ─────
Write-Host "--- V1. intax_backup.json 최신성 확인 ---" -ForegroundColor DarkCyan

if (-not (Test-Path $BackupJson)) {
    Write-Fail "V1. intax_backup.json 파일 없음 ($BackupJson)"
} else {
    $fileAge = (Get-Date) - (Get-Item $BackupJson).LastWriteTime
    $ageMin  = [Math]::Round($fileAge.TotalMinutes, 1)
    if ($fileAge.TotalMinutes -le $MaxAgeMinutes) {
        Write-Pass "V1. intax_backup.json 최신성 OK (최종 수정: ${ageMin}분 전, 기준: ${MaxAgeMinutes}분 이내)"
    } else {
        Write-Fail "V1. intax_backup.json 오래됨 — 최종 수정 ${ageMin}분 전 (기준: ${MaxAgeMinutes}분 이내) — 동기화가 완료되었는지 확인"
    }
}

# ── V2. meta.hometax_synced_at 또는 htxLastSync 필드 존재 ────
Write-Host ""
Write-Host "--- V2. 동기화 타임스탬프 필드 확인 ---" -ForegroundColor DarkCyan

$localData = $null
if (-not (Test-Path $BackupJson)) {
    Write-Fail "V2. intax_backup.json 없음 — 검증 불가"
} else {
    try {
        $localData = [System.IO.File]::ReadAllText($BackupJson, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

        # 스키마 v2.0: meta.hometax_synced_at / 레거시: htxLastSync
        $syncedAt = $null
        if ($localData.meta -ne $null -and $localData.meta.hometax_synced_at -ne $null) {
            $syncedAt = $localData.meta.hometax_synced_at
        } elseif ($localData.htxLastSync -ne $null) {
            $syncedAt = $localData.htxLastSync
        }

        if ($null -eq $syncedAt -or $syncedAt -eq "") {
            Write-Fail "V2. 동기화 타임스탬프 없음 — meta.hometax_synced_at / htxLastSync 필드가 비어 있음"
        } else {
            try {
                $syncTime  = [datetime]::Parse($syncedAt)
                $syncAge   = (Get-Date) - $syncTime
                $syncAgeMin = [Math]::Round($syncAge.TotalMinutes, 1)
                Write-Pass "V2. 동기화 타임스탬프 존재 ($syncedAt, ${syncAgeMin}분 전)"
            } catch {
                Write-Pass "V2. 동기화 타임스탬프 존재 (값: $syncedAt) — 날짜 파싱 주의"
            }
        }
    } catch {
        Write-Fail "V2. intax_backup.json JSON 파싱 실패 — $_"
    }
}

# ── V3. clients 배열 건수 > 이전 건수 ────────────────────────
Write-Host ""
Write-Host "--- V3. 거래처 건수 증가 확인 ---" -ForegroundColor DarkCyan

if ($null -eq $localData) {
    Write-Fail "V3. intax_backup.json 로드 실패 — 건수 비교 불가"
} else {
    $currentCount = @($localData.clients).Count

    if ($PreviousCount -le 0) {
        # 이전 건수를 명시하지 않은 경우: 500건 이상인지만 확인
        if ($currentCount -gt 500) {
            Write-Pass "V3. 거래처 건수 정상 (현재: ${currentCount}건 > 500건) — 이전 건수 비교 생략 (-PreviousCount 옵션 미지정)"
        } else {
            Write-Fail "V3. 거래처 건수 비정상 (현재: ${currentCount}건, 500건 이하) — 동기화 결과 확인 필요"
        }
    } else {
        if ($currentCount -gt $PreviousCount) {
            $diff = $currentCount - $PreviousCount
            Write-Pass "V3. 거래처 건수 증가 확인 (이전: ${PreviousCount}건 → 현재: ${currentCount}건, +${diff}건)"
        } elseif ($currentCount -eq $PreviousCount) {
            Write-Pass "V3. 거래처 건수 변동 없음 (${currentCount}건) — 신규 수임 없을 경우 정상"
        } else {
            $diff = $PreviousCount - $currentCount
            Write-Fail "V3. 거래처 건수 감소 (이전: ${PreviousCount}건 → 현재: ${currentCount}건, -${diff}건) — 병합 로직 확인 필요"
        }
    }
}

# ── V4. GitHub raw URL 데이터 == 로컬 intax_backup.json ──────
Write-Host ""
Write-Host "--- V4. GitHub 업로드 검증 (로컬 vs 원격) ---" -ForegroundColor DarkCyan

if ($null -eq $localData) {
    Write-Fail "V4. 로컬 intax_backup.json 로드 실패 — 비교 불가"
} else {
    try {
        $resp = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop

        if ($resp.StatusCode -ne 200) {
            Write-Fail "V4. GitHub raw URL 접근 실패 (HTTP $($resp.StatusCode))"
        } else {
            try {
                $remoteData   = $resp.Content | ConvertFrom-Json
                $remoteCount  = @($remoteData.clients).Count
                $localCount   = @($localData.clients).Count

                $remoteSynced = $null
                if ($remoteData.meta -ne $null -and $remoteData.meta.hometax_synced_at -ne $null) {
                    $remoteSynced = $remoteData.meta.hometax_synced_at
                } elseif ($remoteData.htxLastSync -ne $null) {
                    $remoteSynced = $remoteData.htxLastSync
                }

                $localSynced = $null
                if ($localData.meta -ne $null -and $localData.meta.hometax_synced_at -ne $null) {
                    $localSynced = $localData.meta.hometax_synced_at
                } elseif ($localData.htxLastSync -ne $null) {
                    $localSynced = $localData.htxLastSync
                }

                $v4Errors = @()

                if ($remoteCount -ne $localCount) {
                    $v4Errors += "거래처 건수 불일치 (로컬: ${localCount}건, 원격: ${remoteCount}건)"
                }

                if ($null -ne $localSynced -and $null -ne $remoteSynced) {
                    if ($localSynced -ne $remoteSynced) {
                        $v4Errors += "동기화 타임스탬프 불일치 (로컬: $localSynced, 원격: $remoteSynced)"
                    }
                }

                if ($v4Errors.Count -eq 0) {
                    Write-Pass "V4. GitHub 원격 데이터 일치 (로컬 ${localCount}건 == 원격 ${remoteCount}건, 타임스탬프 일치)"
                } else {
                    Write-Fail "V4. GitHub 데이터 불일치 — $($v4Errors -join ' | ') — git push가 완료되었는지 확인"
                }

            } catch {
                Write-Fail "V4. GitHub raw URL 응답 JSON 파싱 실패 — $_"
            }
        }
    } catch {
        Write-Fail "V4. GitHub raw URL 접근 실패 — $_ (네트워크 또는 GitHub 접근성 확인)"
    }
}

# ── V5. git log 최신 커밋 메시지 확인 ────────────────────────
Write-Host ""
Write-Host "--- V5. git 커밋 메시지 확인 ---" -ForegroundColor DarkCyan

try {
    $gitLog = & git -C $Root log --oneline -1 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Fail "V5. git log 실행 실패 — $gitLog"
    } else {
        $logStr = "$gitLog"
        if ($logStr -match "홈택스 자동 동기화") {
            Write-Pass "V5. git 최신 커밋에 '홈택스 자동 동기화' 포함 확인 ($logStr)"
        } else {
            Write-Fail "V5. git 최신 커밋에 '홈택스 자동 동기화' 문구 없음 — 최신 커밋: $logStr"
        }
    }
} catch {
    Write-Fail "V5. git 명령 실행 오류 — $_"
}

# ── 최종 요약 ─────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  검증 결과 요약" -ForegroundColor Cyan
Write-Host "================================================================"
Write-Host "  PASS: $($script:PassCount)   FAIL: $($script:FailCount)" -ForegroundColor White
Write-Host "----------------------------------------------------------------"

foreach ($line in $script:Results) {
    if ($line -match "^\[PASS\]") {
        Write-Host "  $line" -ForegroundColor Green
    } else {
        Write-Host "  $line" -ForegroundColor Red
    }
}

Write-Host "================================================================"
Write-Host ""

if ($script:FailCount -gt 0) {
    Write-Host "검증 실패 항목이 있습니다. [FAIL] 항목을 확인하세요." -ForegroundColor Red
    exit 1
} else {
    Write-Host "모든 검증을 통과했습니다. 동기화가 정상적으로 완료되었습니다." -ForegroundColor Green
    exit 0
}
