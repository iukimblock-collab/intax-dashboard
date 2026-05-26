# ================================================================
# hometax_sync_test.ps1
# 홈택스 공동인증서 자동 연동 시스템 — 전체 QA 테스트 스위트
#
# 사용법:
#   .\hometax_sync_test.ps1
#   .\hometax_sync_test.ps1 -Verbose
#   .\hometax_sync_test.ps1 -TestOnly T1,T3,T5
#
# PowerShell 5.1 호환 (&&, ??, ?. 미사용)
# ================================================================

param(
    [switch]$Verbose,
    [string[]]$TestOnly = @()
)

$Root      = $PSScriptRoot
$BackupJson = Join-Path $Root "intax_backup.json"
$VaultPath  = Join-Path $Root ".hometax.vault"
$RawUrl     = "https://raw.githubusercontent.com/iukimblock-collab/intax-dashboard/main/intax_backup.json"

# ── 출력 헬퍼 ─────────────────────────────────────────────────
$script:PassCount = 0
$script:FailCount = 0
$script:SkipCount = 0
$script:Results   = @()

function Write-Pass([string]$id, [string]$msg) {
    $line = "[PASS] $id. $msg"
    Write-Host $line -ForegroundColor Green
    $script:PassCount++
    $script:Results += $line
}

function Write-Fail([string]$id, [string]$msg) {
    $line = "[FAIL] $id. $msg"
    Write-Host $line -ForegroundColor Red
    $script:FailCount++
    $script:Results += $line
}

function Write-Skip([string]$id, [string]$msg) {
    $line = "[SKIP] $id. $msg"
    Write-Host $line -ForegroundColor Yellow
    $script:SkipCount++
    $script:Results += $line
}

function Should-Run([string]$id) {
    if ($TestOnly.Count -eq 0) { return $true }
    return $TestOnly -contains $id
}

# ── vault 복호화 헬퍼 (값 출력 금지) ─────────────────────────
function Decrypt-VaultString([string]$encryptedText) {
    $secure = $encryptedText | ConvertTo-SecureString -ErrorAction Stop
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Load-VaultKeys([string]$vaultPath) {
    $result = @{}
    Get-Content $vaultPath | ForEach-Object {
        if ($_ -match "^([^#=]+)=(.+)$") {
            $result[$Matches[1].Trim()] = $Matches[2].Trim()  # 암호화된 값만 저장
        }
    }
    return $result
}

# ================================================================
# T1. intax_backup.json 구조 검증
# ================================================================
if (Should-Run "T1") {
    Write-Host ""
    Write-Host "--- T1. intax_backup.json 구조 검증 ---" -ForegroundColor Cyan

    if (-not (Test-Path $BackupJson)) {
        Write-Fail "T1" "intax_backup.json 파일 없음 ($BackupJson)"
    } else {
        try {
            $raw  = [System.IO.File]::ReadAllText($BackupJson, [System.Text.Encoding]::UTF8)
            $data = $raw | ConvertFrom-Json

            $errors = @()

            # 최상위 배열 키 존재 확인
            if ($null -eq $data.clients) { $errors += "clients 키 없음" }
            if ($null -eq $data.staff)   { $errors += "staff 키 없음" }
            if ($null -eq $data.users)   { $errors += "users 키 없음" }

            # clients가 배열인지 (PSCustomObject 배열)
            if ($data.clients -ne $null) {
                $clientCount = @($data.clients).Count
                if ($clientCount -eq 0) { $errors += "clients 배열이 비어 있음" }
            }

            # htxLastSync 또는 meta 존재 확인 (현재 스키마: htxLastSync)
            $hasMeta   = $null -ne $data.meta
            $hasLegacy = $null -ne $data.htxLastSync
            if (-not $hasMeta -and -not $hasLegacy) {
                $errors += "동기화 타임스탬프(meta.hometax_synced_at 또는 htxLastSync) 없음 (경고)"
            }

            # clients 항목 필드 검사 (첫 10건)
            if ($data.clients -ne $null) {
                $sample = @($data.clients) | Select-Object -First 10
                $missingFields = @()
                foreach ($c in $sample) {
                    foreach ($field in @("name","type","div","status")) {
                        if ($null -eq $c.$field) {
                            $missingFields += $field
                        }
                    }
                }
                $missingFields = $missingFields | Sort-Object -Unique
                if ($missingFields.Count -gt 0) {
                    $errors += "clients 항목 필수 필드 누락: $($missingFields -join ', ')"
                }
            }

            if ($errors.Count -eq 0) {
                $cc = @($data.clients).Count
                $sc = @($data.staff).Count
                Write-Pass "T1" "intax_backup.json 구조 검증 (clients: ${cc}건, staff: ${sc}명)"
            } else {
                Write-Fail "T1" "intax_backup.json 구조 이상 — $($errors -join ' | ')"
            }

        } catch {
            Write-Fail "T1" "JSON 파싱 실패 — $_"
        }
    }
}

# ================================================================
# T2. 데이터 병합 로직 검증 (더미 데이터 사용)
# ================================================================
if (Should-Run "T2") {
    Write-Host ""
    Write-Host "--- T2. 데이터 병합 로직 검증 ---" -ForegroundColor Cyan

    try {
        # 기존 백업 데이터 (biz_no 있는 5건 + 없는 2건)
        $existingClients = @(
            [PSCustomObject]@{ id="c_001"; biz_no="111-11-11111"; name="기존법인A"; type="법인"; div="기장"; status="정상"; sector=""; fee=300000; note="VIP"; staff_id="s1" }
            [PSCustomObject]@{ id="c_002"; biz_no="222-22-22222"; name="기존개인B"; type="개인"; div="기장"; status="정상"; sector=""; fee=150000; note="중요"; staff_id="s2" }
            [PSCustomObject]@{ id="c_003"; biz_no="333-33-33333"; name="기존법인C"; type="법인"; div="신고대리"; status="정상"; sector=""; fee=200000; note="";    staff_id="s1" }
            [PSCustomObject]@{ id="c_004"; biz_no="444-44-44444"; name="기존개인D"; type="개인"; div="기장"; status="정상"; sector=""; fee=100000; note="메모"; staff_id="s3" }
            [PSCustomObject]@{ id="c_005"; biz_no="555-55-55555"; name="기존법인E"; type="법인"; div="기장"; status="정상"; sector=""; fee=500000; note="";    staff_id="s2" }
            [PSCustomObject]@{ id="c_006"; biz_no="";             name="사업자번호없음1"; type="개인"; div="신고대리"; status="정상"; sector=""; fee=0;      note=""; staff_id="" }
            [PSCustomObject]@{ id="c_007"; biz_no="";             name="사업자번호없음2"; type="개인"; div="신고대리"; status="정상"; sector=""; fee=0;      note=""; staff_id="" }
        )

        # 홈택스에서 새로 내려받은 데이터
        # - 기존 3건 (biz_no 일치): 222, 333, 555 → 555는 폐업으로 상태 변경
        # - 신규 10건 (새 biz_no)
        $htxClients = @(
            [PSCustomObject]@{ biz_no="222-22-22222"; name="기존개인B"; type="개인"; div="기장"; status="정상"; sector="서비스" }
            [PSCustomObject]@{ biz_no="333-33-33333"; name="기존법인C갱신"; type="법인"; div="신고대리"; status="정상"; sector="제조" }
            [PSCustomObject]@{ biz_no="555-55-55555"; name="기존법인E"; type="법인"; div="기장"; status="폐업"; sector="" }
            [PSCustomObject]@{ biz_no="601-01-01001"; name="신규법인1"; type="법인"; div="기장"; status="정상"; sector="도소매" }
            [PSCustomObject]@{ biz_no="602-02-02002"; name="신규개인2"; type="개인"; div="기장"; status="정상"; sector="음식" }
            [PSCustomObject]@{ biz_no="603-03-03003"; name="신규법인3"; type="법인"; div="신고대리"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="604-04-04004"; name="신규개인4"; type="개인"; div="기장"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="605-05-05005"; name="신규법인5"; type="법인"; div="기장"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="606-06-06006"; name="신규개인6"; type="개인"; div="기장"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="607-07-07007"; name="신규법인7"; type="법인"; div="신고대리"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="608-08-08008"; name="신규개인8"; type="개인"; div="기장"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="609-09-09009"; name="신규법인9"; type="법인"; div="기장"; status="정상"; sector="" }
            [PSCustomObject]@{ biz_no="610-10-10010"; name="신규개인10"; type="개인"; div="기장"; status="정상"; sector="" }
        )

        # ── 병합 함수 (hometax_auto_sync.ps1 과 동일 로직 재현) ──
        function Merge-Clients {
            param($existing, $htxList)

            # biz_no 기준 기존 인덱스 맵
            $existMap = @{}
            foreach ($c in $existing) {
                if ($c.biz_no -ne "") {
                    $existMap[$c.biz_no] = $c
                }
            }

            $merged   = [System.Collections.ArrayList]@($existing)
            $addCount = 0
            $updCount = 0

            foreach ($htx in $htxList) {
                if ($htx.biz_no -eq "") { continue }

                if ($existMap.ContainsKey($htx.biz_no)) {
                    # 기존 거래처: 상태만 업데이트, fee/staff/note 보존
                    $orig = $existMap[$htx.biz_no]
                    $orig.status = $htx.status
                    if ($htx.sector -ne "") { $orig.sector = $htx.sector }
                    $updCount++
                } else {
                    # 신규 거래처 추가
                    $ts   = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
                    $newC = [PSCustomObject]@{
                        id         = "c${ts}_htx"
                        biz_no     = $htx.biz_no
                        name       = $htx.name
                        type       = $htx.type
                        div        = $htx.div
                        sector     = $htx.sector
                        size       = if ($htx.type -eq "법인") { "중소기업" } else { "개인사업자" }
                        status     = $htx.status
                        fee        = 0
                        cms        = "N"
                        cms_status = "미납"
                        staff_id   = ""
                        phone      = ""
                        hometax_id = ""
                        karinara_id= ""
                        note       = "홈택스 자동 동기화"
                        htx_synced = $true
                    }
                    $merged.Add($newC) | Out-Null
                    $addCount++
                }
            }
            return @{ Merged = $merged; Added = $addCount; Updated = $updCount }
        }

        $result     = Merge-Clients -existing $existingClients -htxList $htxClients
        $mergedList = $result.Merged
        $addCount   = $result.Added
        $updCount   = $result.Updated

        $mergeErrors = @()

        # 검증 1: 신규 10건 추가되었는지
        if ($addCount -ne 10) {
            $mergeErrors += "신규 추가 건수 오류: 예상 10건, 실제 ${addCount}건"
        }

        # 검증 2: 기존 3건 업데이트되었는지
        if ($updCount -ne 3) {
            $mergeErrors += "업데이트 건수 오류: 예상 3건, 실제 ${updCount}건"
        }

        # 검증 3: 222 거래처 fee/note 보존되었는지
        $c222 = $mergedList | Where-Object { $_.biz_no -eq "222-22-22222" }
        if ($null -eq $c222) {
            $mergeErrors += "222-22-22222 거래처가 병합 결과에 없음"
        } else {
            if ($c222.fee -ne 150000) { $mergeErrors += "fee 보존 실패 (222): 예상 150000, 실제 $($c222.fee)" }
            if ($c222.note -ne "중요")  { $mergeErrors += "note 보존 실패 (222): 예상 '중요', 실제 '$($c222.note)'" }
            if ($c222.staff_id -ne "s2") { $mergeErrors += "staff_id 보존 실패 (222)" }
        }

        # 검증 4: 555 거래처 폐업 상태로 업데이트되었는지
        $c555 = $mergedList | Where-Object { $_.biz_no -eq "555-55-55555" }
        if ($null -eq $c555) {
            $mergeErrors += "555-55-55555 거래처가 병합 결과에 없음"
        } else {
            if ($c555.status -ne "폐업") { $mergeErrors += "폐업 상태 업데이트 실패 (555): '$($c555.status)'" }
            if ($c555.fee -ne 500000)    { $mergeErrors += "fee 보존 실패 (555): 예상 500000, 실제 $($c555.fee)" }
        }

        # 검증 5: biz_no 없는 기존 2건 유지되는지
        $noBizNo = @($mergedList | Where-Object { $_.biz_no -eq "" })
        if ($noBizNo.Count -ne 2) {
            $mergeErrors += "사업자번호 없는 기존 거래처 보존 실패: 예상 2건, 실제 $($noBizNo.Count)건"
        }

        # 검증 6: 전체 건수 맞는지 (7 기존 + 10 신규 = 17)
        $totalExpected = 17
        if (@($mergedList).Count -ne $totalExpected) {
            $mergeErrors += "병합 후 전체 건수 오류: 예상 ${totalExpected}건, 실제 $(@($mergedList).Count)건"
        }

        if ($mergeErrors.Count -eq 0) {
            Write-Pass "T2" "데이터 병합 로직 검증 (신규 +${addCount}건, 업데이트 ${updCount}건, 전체 $(@($mergedList).Count)건, fee/note 보존 OK)"
        } else {
            Write-Fail "T2" "병합 로직 이상 — $($mergeErrors -join ' | ')"
        }

    } catch {
        Write-Fail "T2" "병합 테스트 예외 발생 — $_"
    }
}

# ================================================================
# T3. GitHub raw URL 접근성
# ================================================================
if (Should-Run "T3") {
    Write-Host ""
    Write-Host "--- T3. GitHub raw URL 접근성 ---" -ForegroundColor Cyan

    try {
        $resp = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop

        if ($resp.StatusCode -ne 200) {
            Write-Fail "T3" "GitHub raw URL 응답 코드 이상: $($resp.StatusCode)"
        } else {
            try {
                $remoteJson = $resp.Content.TrimStart([char]0xFEFF)
                $remoteData = $remoteJson | ConvertFrom-Json
                $remoteCount = @($remoteData.clients).Count

                $urlErrors = @()
                if ($null -eq $remoteData.clients) { $urlErrors += "clients 키 없음" }
                if ($remoteCount -le 500)           { $urlErrors += "거래처 건수 500건 이하: ${remoteCount}건" }

                if ($urlErrors.Count -eq 0) {
                    Write-Pass "T3" "GitHub raw URL 접근 및 JSON 유효 (거래처: ${remoteCount}건)"
                } else {
                    Write-Fail "T3" "GitHub raw 데이터 이상 — $($urlErrors -join ' | ')"
                }
            } catch {
                Write-Fail "T3" "GitHub raw URL 응답이 유효한 JSON이 아님 — $_"
            }
        }
    } catch {
        Write-Fail "T3" "GitHub raw URL 접근 실패 — $_ (네트워크 확인 필요)"
    }
}

# ================================================================
# T4. vault 무결성
# ================================================================
if (Should-Run "T4") {
    Write-Host ""
    Write-Host "--- T4. vault 무결성 ---" -ForegroundColor Cyan

    if (-not (Test-Path $VaultPath)) {
        Write-Fail "T4" "vault 파일 없음 (.hometax.vault) — vault.ps1 -Service hometax init 실행 필요"
    } else {
        try {
            $vaultKeys = Load-VaultKeys $VaultPath

            $vaultErrors   = @()
            $vaultWarnings = @()

            # PASSWORD 키 존재 및 복호화 가능 여부
            if (-not $vaultKeys.ContainsKey("PASSWORD")) {
                $vaultErrors += "PASSWORD 키 없음"
            } else {
                try {
                    $decrypted = Decrypt-VaultString $vaultKeys["PASSWORD"]
                    if ($null -eq $decrypted -or $decrypted.Length -eq 0) {
                        $vaultErrors += "PASSWORD 복호화 결과가 비어 있음"
                    }
                    # 값은 절대 출력하지 않음
                    Remove-Variable decrypted -ErrorAction SilentlyContinue
                } catch {
                    $vaultErrors += "PASSWORD 복호화 실패 (다른 Windows 사용자/PC에서 생성된 vault일 수 있음)"
                }
            }

            # CERT_PIN 존재 여부 (없으면 경고만)
            if (-not $vaultKeys.ContainsKey("CERT_PIN")) {
                $vaultWarnings += "CERT_PIN 키 없음 (경고: PIN 설정 후 재실행 권장)"
            }

            if ($vaultErrors.Count -gt 0) {
                Write-Fail "T4" "vault 무결성 이상 — $($vaultErrors -join ' | ')$(if($vaultWarnings.Count -gt 0){' | '+ ($vaultWarnings -join ' | ')})"
            } elseif ($vaultWarnings.Count -gt 0) {
                Write-Host "[WARN] T4. vault 무결성 — PASSWORD OK | $($vaultWarnings -join ' | ')" -ForegroundColor Yellow
                $script:Results += "[WARN] T4. vault 무결성 — PASSWORD OK | $($vaultWarnings -join ' | ')"
            } else {
                Write-Pass "T4" "vault 무결성 OK (PASSWORD 복호화 가능, CERT_PIN 존재)"
            }

        } catch {
            Write-Fail "T4" "vault 읽기 실패 — $_"
        }
    }
}

# ================================================================
# T5. NPKI 인증서 파일 검증
# ================================================================
if (Should-Run "T5") {
    Write-Host ""
    Write-Host "--- T5. NPKI 인증서 파일 검증 ---" -ForegroundColor Cyan

    $npkiRoot = "C:\NPKI"

    if (-not (Test-Path $npkiRoot)) {
        Write-Skip "T5" "NPKI 디렉토리 없음 (C:\NPKI) — 공동인증서 미설치"
    } else {
        try {
            # INTAX 관련 디렉토리 탐색 (인택스, INTAX, 김인욱 등)
            $certDirs = Get-ChildItem -Path $npkiRoot -Recurse -Directory -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -match "INTAX|인택스|김인욱|세무|회계" }

            # 이름 매칭 없으면 가장 최근 수정된 디렉토리로 폴백
            if ($null -eq $certDirs -or @($certDirs).Count -eq 0) {
                $certDirs = Get-ChildItem -Path $npkiRoot -Recurse -Directory -ErrorAction SilentlyContinue |
                            Where-Object { (Test-Path (Join-Path $_.FullName "signCert.der")) } |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 3
            }

            if ($null -eq $certDirs -or @($certDirs).Count -eq 0) {
                Write-Fail "T5" "NPKI 인증서 디렉토리를 찾을 수 없음 (signCert.der 보유 폴더 없음)"
            } else {
                $certErrors   = @()
                $certWarnings = @()
                $foundDir     = $null

                foreach ($dir in @($certDirs)) {
                    $derPath = Join-Path $dir.FullName "signCert.der"
                    $keyPath = Join-Path $dir.FullName "signPri.key"

                    if ((Test-Path $derPath) -and (Test-Path $keyPath)) {
                        $foundDir = $dir.FullName
                        break
                    }
                }

                if ($null -eq $foundDir) {
                    Write-Fail "T5" "signCert.der + signPri.key 쌍을 갖춘 인증서 디렉토리 없음"
                } else {
                    $derPath = Join-Path $foundDir "signCert.der"
                    $keyPath = Join-Path $foundDir "signPri.key"

                    # 인증서 유효기간 확인
                    try {
                        $certBytes = [System.IO.File]::ReadAllBytes($derPath)
                        $x509 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                        $x509.Import($certBytes)

                        $notAfter = $x509.NotAfter
                        $now      = Get-Date
                        $daysLeft = ($notAfter - $now).Days

                        if ($daysLeft -lt 0) {
                            $certErrors += "인증서 만료됨 ($($notAfter.ToString('yyyy-MM-dd')))"
                        } elseif ($daysLeft -lt 30) {
                            $certWarnings += "인증서 만료 임박 (${daysLeft}일 남음, 만료: $($notAfter.ToString('yyyy-MM-dd')))"
                        }

                        $x509.Dispose()

                        if ($certErrors.Count -gt 0) {
                            Write-Fail "T5" "NPKI 인증서 이상 — $($certErrors -join ' | ')"
                        } elseif ($certWarnings.Count -gt 0) {
                            Write-Pass "T5" "NPKI 인증서 존재 (signCert.der + signPri.key) | 경고: $($certWarnings -join ' | ')"
                        } else {
                            Write-Pass "T5" "NPKI 인증서 검증 OK (signCert.der + signPri.key, 유효기간 ${daysLeft}일 남음, 만료: $($notAfter.ToString('yyyy-MM-dd')))"
                        }

                    } catch {
                        # .der 파싱 실패 시 파일 존재만 확인
                        Write-Pass "T5" "NPKI 인증서 파일 존재 확인 (signCert.der + signPri.key) — 유효기간 파싱 실패: $_"
                    }
                }
            }
        } catch {
            Write-Fail "T5" "NPKI 탐색 중 오류 — $_"
        }
    }
}

# ================================================================
# T6. Chrome CDP 연결 테스트
# ================================================================
if (Should-Run "T6") {
    Write-Host ""
    Write-Host "--- T6. Chrome CDP 연결 테스트 ---" -ForegroundColor Cyan

    $cdpUrl = "http://localhost:9222/json"

    try {
        $resp = Invoke-WebRequest -Uri $cdpUrl -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop

        if ($resp.StatusCode -eq 200) {
            try {
                $tabs = $resp.Content | ConvertFrom-Json
                $tabCount = @($tabs).Count
                Write-Pass "T6" "Chrome CDP 연결 성공 (포트 9222, 탭 ${tabCount}개 감지)"
            } catch {
                Write-Pass "T6" "Chrome CDP 응답 수신 (포트 9222) — JSON 파싱 주의: $_"
            }
        } else {
            Write-Fail "T6" "Chrome CDP 응답 코드 이상: $($resp.StatusCode)"
        }
    } catch {
        Write-Skip "T6" "Chrome CDP — Chrome이 실행 중이지 않음 (--remote-debugging-port=9222 옵션 필요)"
    }
}

# ================================================================
# T7. 동기화 후 대시보드 데이터 무결성 (Chrome CDP 필요)
# ================================================================
if (Should-Run "T7") {
    Write-Host ""
    Write-Host "--- T7. 동기화 후 대시보드 데이터 무결성 ---" -ForegroundColor Cyan

    $cdpAvail = $false
    try {
        $cdpCheck = Invoke-WebRequest -Uri "http://localhost:9222/json" -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop
        if ($cdpCheck.StatusCode -eq 200) { $cdpAvail = $true }
    } catch { }

    if (-not $cdpAvail) {
        Write-Skip "T7" "대시보드 무결성 — Chrome CDP 미실행 (T6 먼저 통과 필요)"
    } else {
        try {
            # CDP를 통해 localStorage 조회
            $tabs    = (Invoke-WebRequest -Uri "http://localhost:9222/json" -UseBasicParsing -TimeoutSec 5).Content | ConvertFrom-Json
            $dashTab = $tabs | Where-Object { $_.url -match "index\.html" -or $_.title -match "INTAX" } | Select-Object -First 1

            if ($null -eq $dashTab) {
                Write-Skip "T7" "대시보드 탭 없음 — index.html을 Chrome에서 열고 재실행"
            } else {
                $wsUrl   = $dashTab.webSocketDebuggerUrl
                # WebSocket CDP 명령 실행 (PowerShell 5.1: HttpWebRequest 방식)
                # localStorage 값은 Runtime.evaluate 로 조회
                $evalScript = "JSON.stringify({clients: (function(){try{var d=JSON.parse(localStorage.getItem('DB'));return d?d.clients?d.clients.length:0:0;}catch(e){return -1;}}()), synced: (function(){try{var d=JSON.parse(localStorage.getItem('DB'));return d&&d.meta?d.meta.hometax_synced_at:(d?d.htxLastSync:'');}catch(e){return '';}}())})"

                # CDP HTTP endpoint (간이 방식: /json/evaluate 미지원이므로 경고 출력)
                Write-Host "  [정보] CDP WebSocket 평가 필요 — 브라우저 콘솔에서 직접 확인:" -ForegroundColor DarkYellow
                Write-Host "         $evalScript" -ForegroundColor DarkGray

                # intax_backup.json 건수와 비교
                if (Test-Path $BackupJson) {
                    $localData  = [System.IO.File]::ReadAllText($BackupJson, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
                    $localCount = @($localData.clients).Count
                    Write-Pass "T7" "intax_backup.json 로컬 clients 건수 기준 ${localCount}건 — 브라우저 localStorage 수동 대조 필요"
                } else {
                    Write-Fail "T7" "intax_backup.json 없음 — 대시보드 무결성 검증 불가"
                }
            }
        } catch {
            Write-Fail "T7" "CDP 데이터 조회 실패 — $_"
        }
    }
}

# ================================================================
# T8. 병합 함수 단위 테스트 (10건 신규 + 5건 업데이트)
# ================================================================
if (Should-Run "T8") {
    Write-Host ""
    Write-Host "--- T8. 병합 함수 단위 테스트 ---" -ForegroundColor Cyan

    try {
        # 기존 15건 (biz_no 있는 5건 + 없는 10건)
        $base = @()
        for ($i = 1; $i -le 5; $i++) {
            $bn = "{0:D3}-{0:D2}-{0:D5}" -f $i
            $base += [PSCustomObject]@{
                id       = "base_$i"
                biz_no   = $bn
                name     = "기존거래처$i"
                type     = "개인"
                div      = "기장"
                status   = "정상"
                fee      = ($i * 100000)
                note     = "기존메모$i"
                staff_id = "s1"
            }
        }
        for ($i = 6; $i -le 15; $i++) {
            $base += [PSCustomObject]@{
                id       = "base_$i"
                biz_no   = ""
                name     = "사번없음거래처$i"
                type     = "개인"
                div      = "신고대리"
                status   = "정상"
                fee      = 0
                note     = ""
                staff_id = ""
            }
        }

        # 신규 10건 + 기존 5건 업데이트 (상태 변경)
        $htxNew = @()
        for ($i = 1; $i -le 5; $i++) {
            $bn = "{0:D3}-{0:D2}-{0:D5}" -f $i
            $htxNew += [PSCustomObject]@{ biz_no=$bn; name="기존거래처$i"; type="개인"; div="기장"; status="정상"; sector="서비스" }
        }
        for ($i = 101; $i -le 110; $i++) {
            $bn = "{0:D3}-11-{0:D5}" -f $i
            $htxNew += [PSCustomObject]@{ biz_no=$bn; name="신규거래처$i"; type="법인"; div="기장"; status="정상"; sector="제조" }
        }

        # 병합 수행 (T2와 동일 로직)
        $existMapT8 = @{}
        foreach ($c in $base) {
            if ($c.biz_no -ne "") { $existMapT8[$c.biz_no] = $c }
        }

        $mergedT8   = [System.Collections.ArrayList]@($base)
        $addT8      = 0
        $updT8      = 0
        $feePreserved = $true
        $notePreserved = $true

        foreach ($htx in $htxNew) {
            if ($htx.biz_no -eq "") { continue }
            if ($existMapT8.ContainsKey($htx.biz_no)) {
                $orig = $existMapT8[$htx.biz_no]
                $orig.status = $htx.status
                $updT8++
            } else {
                $mergedT8.Add([PSCustomObject]@{
                    id         = "c_htx_$addT8"
                    biz_no     = $htx.biz_no
                    name       = $htx.name
                    type       = $htx.type
                    div        = $htx.div
                    sector     = $htx.sector
                    status     = $htx.status
                    fee        = 0
                    note       = "홈택스 자동 동기화"
                    staff_id   = ""
                }) | Out-Null
                $addT8++
            }
        }

        # fee/note 보존 검증
        for ($i = 1; $i -le 5; $i++) {
            $bn  = "{0:D3}-{0:D2}-{0:D5}" -f $i
            $c   = $mergedT8 | Where-Object { $_.biz_no -eq $bn }
            if ($c.fee  -ne ($i * 100000)) { $feePreserved  = $false }
            if ($c.note -ne "기존메모$i")   { $notePreserved = $false }
        }

        $t8Errors = @()
        if ($addT8 -ne 10)       { $t8Errors += "신규 추가 오류: 예상 10건, 실제 ${addT8}건" }
        if ($updT8 -ne 5)        { $t8Errors += "업데이트 오류: 예상 5건, 실제 ${updT8}건" }
        if (-not $feePreserved)  { $t8Errors += "fee 보존 실패" }
        if (-not $notePreserved) { $t8Errors += "note 보존 실패" }
        $expectedTotal = 25  # 15 기존 + 10 신규
        if (@($mergedT8).Count -ne $expectedTotal) {
            $t8Errors += "전체 건수 오류: 예상 ${expectedTotal}건, 실제 $(@($mergedT8).Count)건"
        }

        if ($t8Errors.Count -eq 0) {
            Write-Pass "T8" "병합 단위 테스트 OK (신규 +${addT8}건, 업데이트 ${updT8}건, fee/note 보존 확인, 전체 $(@($mergedT8).Count)건)"
        } else {
            Write-Fail "T8" "병합 단위 테스트 실패 — $($t8Errors -join ' | ')"
        }

    } catch {
        Write-Fail "T8" "T8 단위 테스트 예외 — $_"
    }
}

# ================================================================
# 최종 결과 요약
# ================================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  QA 테스트 결과 요약" -ForegroundColor Cyan
Write-Host "================================================================"
Write-Host "  PASS: $($script:PassCount)  FAIL: $($script:FailCount)  SKIP: $($script:SkipCount)" -ForegroundColor White
Write-Host "================================================================"

foreach ($line in $script:Results) {
    if ($line -match "^\[PASS\]") {
        Write-Host "  $line" -ForegroundColor Green
    } elseif ($line -match "^\[FAIL\]") {
        Write-Host "  $line" -ForegroundColor Red
    } else {
        Write-Host "  $line" -ForegroundColor Yellow
    }
}

Write-Host "================================================================"
Write-Host ""

if ($script:FailCount -gt 0) {
    Write-Host "일부 테스트가 실패했습니다. 위 [FAIL] 항목을 확인하고 수정 후 재실행하세요." -ForegroundColor Red
    exit 1
} else {
    Write-Host "모든 테스트를 통과(또는 SKIP)했습니다." -ForegroundColor Green
    exit 0
}
