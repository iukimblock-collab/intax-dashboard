# ================================================================
# hometax_auto_sync.ps1
# 홈택스 공동인증서 자동 연동 & 백업 동기화 오케스트레이터
#
# [사용법]
#   .\hometax_auto_sync.ps1                     # 전체 자동화 실행
#   .\hometax_auto_sync.ps1 -SkipBrowser        # 브라우저 자동화 건너뜀 (다운로드 파일 수동 지정 시)
#   .\hometax_auto_sync.ps1 -InputFile "파일.xlsx" # 기존 엑셀로 병합만 실행
#   .\hometax_auto_sync.ps1 -NoPush             # git push 건너뜀
#
# [사전 조건]
#   1. vault.ps1로 .hometax.vault 생성
#      .env(hometax).txt 파일에 아래와 같이 작성 후 실행:
#        PASSWORD=홈택스비밀번호
#        CERT_PIN=공동인증서PIN번호
#      .\vault.ps1 -Service hometax init
#   2. Chrome이 설치되어 있어야 함 (CDP 원격 디버깅 사용)
#   3. hometax_import.ps1이 같은 디렉토리에 있어야 함
#
# [주의사항]
#   - CERT_PIN은 vault에서만 읽어 메모리에 보관. 로그/파일에 절대 노출 금지.
#   - 홈택스 페이지 구조가 변경되면 셀렉터를 수정해야 합니다.
#   - PowerShell 5.1 호환 (??/&&/2>&1 불가)
# ================================================================

param(
    [switch]$SkipBrowser,
    [string]$InputFile   = "",
    [switch]$NoPush,
    [switch]$Verbose
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$Root         = $PSScriptRoot
$VaultScript  = Join-Path $Root "vault.ps1"
$ImportScript = Join-Path $Root "hometax_import.ps1"
$BackupFile   = Join-Path $Root "intax_backup.json"
$HtxDataFile  = Join-Path $Root "hometax_data.json"
$ChromePath   = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$DownloadDir  = Join-Path $env:USERPROFILE "Downloads"
$DebugPort    = 9222

# ─── 로그 헬퍼 ──────────────────────────────────────
function Log([string]$msg, [string]$color="White") {
    Write-Host "  $msg" -ForegroundColor $color
}
function LogOK([string]$msg)   { Write-Host "  [OK]  $msg" -ForegroundColor Green }
function LogWarn([string]$msg) { Write-Host "  [!!]  $msg" -ForegroundColor Yellow }
function LogErr([string]$msg)  { Write-Host "  [ERR] $msg" -ForegroundColor Red }
function LogStep([string]$msg) { Write-Host "`n── $msg" -ForegroundColor Cyan }

# ─── STEP 0: Vault 로드 ──────────────────────────────
LogStep "0. Vault 자격증명 로드"

if (-not (Test-Path $VaultScript)) {
    LogErr "vault.ps1을 찾을 수 없습니다: $VaultScript"
    exit 1
}

$HOMETAX_PW   = $null
$CERT_PIN     = $null

try {
    $HOMETAX_PW = & $VaultScript -Service hometax get PASSWORD 2>$null
    if (-not $HOMETAX_PW) { throw "PASSWORD 키가 비어있습니다." }
    LogOK "PASSWORD 로드 완료"
} catch {
    LogErr "홈택스 비밀번호를 vault에서 읽지 못했습니다: $_"
    Log "실행: .\vault.ps1 -Service hometax init  (먼저 .env(hometax).txt에 PASSWORD= 작성)" Yellow
    exit 1
}

try {
    $CERT_PIN = & $VaultScript -Service hometax get CERT_PIN 2>$null
    if (-not $CERT_PIN) { throw "CERT_PIN 키가 비어있습니다." }
    LogOK "CERT_PIN 로드 완료"
} catch {
    LogErr "공동인증서 PIN을 vault에서 읽지 못했습니다: $_"
    Log ".env(hometax).txt에 CERT_PIN=<PIN번호> 추가 후 .\vault.ps1 -Service hometax init 실행" Yellow
    exit 1
}

# ─── STEP 1: NPKI 인증서 탐색 ────────────────────────
LogStep "1. NPKI 공동인증서 탐색"

$NpkiRoot  = "C:\NPKI"
$IntaxCert = $null
$_now      = Get-Date

if (Test-Path $NpkiRoot) {
    $certDirs = Get-ChildItem -Path $NpkiRoot -Recurse -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "인택스|INTAX|97769" }

    # 파일 존재 + 유효기간 검사 후 만료일 가장 늦은 인증서 선택
    $validCerts = @()
    foreach ($dir in $certDirs) {
        $signCert = Join-Path $dir.FullName "signCert.der"
        $signKey  = Join-Path $dir.FullName "signPri.key"
        if (-not ((Test-Path $signCert) -and (Test-Path $signKey))) { continue }
        try {
            $x509 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $x509.Import([System.IO.File]::ReadAllBytes($signCert))
            if ($x509.NotAfter -gt $_now) {
                $validCerts += [PSCustomObject]@{
                    Dir      = $dir.FullName
                    CertFile = $signCert
                    KeyFile  = $signKey
                    Name     = $dir.Name
                    NotAfter = $x509.NotAfter
                }
            }
            $x509.Dispose()
        } catch { }
    }

    if ($validCerts.Count -gt 0) {
        $best = $validCerts | Sort-Object NotAfter -Descending | Select-Object -First 1
        $IntaxCert = @{
            Dir      = $best.Dir
            CertFile = $best.CertFile
            KeyFile  = $best.KeyFile
            Name     = $best.Name
        }
        $daysLeft = ($best.NotAfter - $_now).Days
        LogOK "인택스 인증서 발견: $($best.Name)"
        LogOK "  경로: $($best.Dir)"
        LogOK "  만료: $($best.NotAfter.ToString('yyyy-MM-dd')) (${daysLeft}일 남음)"
    }
}

if (-not $IntaxCert) {
    LogWarn "C:\NPKI에서 유효한 인택스 인증서를 찾지 못했습니다."
    LogWarn "인증서가 만료되었거나 경로에 '인택스' 또는 '97769'가 포함되지 않을 수 있습니다."
    if (-not $SkipBrowser) {
        Log "브라우저 자동화 시 인증서를 수동으로 선택해야 할 수 있습니다." Yellow
    }
}

# ─── STEP 2: 입력 파일이 이미 있으면 브라우저 건너뜀 ──
if ($InputFile -ne "" -and (Test-Path $InputFile)) {
    LogStep "2. 수동 파일 입력 모드 (브라우저 자동화 건너뜀)"
    LogOK "입력 파일: $InputFile"
    $SkipBrowser = $true
}

# ─── STEP 3: Chrome CDP 자동화 ────────────────────────
if (-not $SkipBrowser) {
    LogStep "2. Chrome 원격 디버깅 시작"

    # 기존 CDP 포트 사용 중인지 확인
    $portInUse = $false
    try {
        $resp = Invoke-WebRequest -Uri "http://localhost:$DebugPort/json" -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue
        if ($resp.StatusCode -eq 200) { $portInUse = $true }
    } catch {}

    if (-not $portInUse) {
        if (-not (Test-Path $ChromePath)) {
            # 대안 경로 시도
            $altPaths = @(
                "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
                "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            )
            foreach ($alt in $altPaths) {
                if (Test-Path $alt) { $ChromePath = $alt; break }
            }
        }
        if (-not (Test-Path $ChromePath)) {
            LogErr "Chrome을 찾을 수 없습니다. Chrome 설치 경로를 확인하세요."
            exit 1
        }

        Log "Chrome 원격 디버깅 시작 (포트: $DebugPort)..."
        $chromeArgs = "--remote-debugging-port=$DebugPort --no-first-run --no-default-browser-check --user-data-dir=`"$env:TEMP\intax-chrome-cdp`""
        Start-Process -FilePath $ChromePath -ArgumentList $chromeArgs
        Start-Sleep -Seconds 5

        # CDP 포트 대기 (최대 60초, 10초 타임아웃으로 체크)
        $waited = 0
        while ($waited -lt 60) {
            try {
                $r = Invoke-WebRequest -Uri "http://localhost:$DebugPort/json" -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue
                if ($r.StatusCode -eq 200) { break }
            } catch {}
            Start-Sleep -Seconds 2
            $waited += 2
        }
        if ($waited -ge 60) {
            LogErr "Chrome CDP 포트($DebugPort) 연결 실패. Chrome이 시작되지 않았습니다."
            exit 1
        }
        LogOK "Chrome CDP 준비 완료"
    } else {
        LogOK "기존 Chrome CDP 포트 사용 ($DebugPort)"
    }

    # ── CDP WebSocket 헬퍼 함수 ─────────────────────────
    function Invoke-CDP {
        param(
            [string]$wsUrl,
            [hashtable]$Message,
            [int]$TimeoutSec = 30
        )
        $ws = $null
        try {
            $ws = New-Object System.Net.WebSockets.ClientWebSocket
            $cts = New-Object System.Threading.CancellationTokenSource
            $connectTask = $ws.ConnectAsync([uri]$wsUrl, $cts.Token)
            $connectTask.Wait(5000) | Out-Null
            if ($ws.State -ne [System.Net.WebSockets.WebSocketState]::Open) {
                throw "WebSocket 연결 실패"
            }

            $json = $Message | ConvertTo-Json -Compress -Depth 10
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
            $seg = New-Object System.ArraySegment[byte] -ArgumentList (,$bytes)
            $ws.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait(5000) | Out-Null

            # 응답 수신 (64KB 청크 루프)
            $buf = New-Object byte[] 65536
            $received = New-Object System.Text.StringBuilder
            do {
                $seg2 = New-Object System.ArraySegment[byte] -ArgumentList (,$buf)
                $recvTask = $ws.ReceiveAsync($seg2, $cts.Token)
                $recvTask.Wait($TimeoutSec * 1000) | Out-Null
                $result = $recvTask.Result
                $received.Append([System.Text.Encoding]::UTF8.GetString($buf, 0, $result.Count)) | Out-Null
            } while (-not $result.EndOfMessage)

            return $received.ToString() | ConvertFrom-Json
        } finally {
            if ($ws -and $ws.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                try { $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "", [System.Threading.CancellationToken]::None).Wait(3000) | Out-Null } catch {}
            }
            if ($ws) { $ws.Dispose() }
        }
    }

    function Invoke-CDPEval {
        param([string]$wsUrl, [string]$Expression, [int]$TimeoutSec=30)
        $msg = @{
            id     = 1
            method = "Runtime.evaluate"
            params = @{
                expression            = $Expression
                returnByValue         = $true
                awaitPromise          = $true
            }
        }
        $resp = Invoke-CDP -wsUrl $wsUrl -Message $msg -TimeoutSec $TimeoutSec
        return $resp
    }

    function Wait-PageReady {
        param([string]$wsUrl, [int]$MaxWaitSec=30)
        $elapsed = 0
        while ($elapsed -lt $MaxWaitSec) {
            try {
                $resp = Invoke-CDPEval -wsUrl $wsUrl -Expression "document.readyState"
                if ($resp.result.result.value -eq "complete") { return $true }
            } catch {}
            Start-Sleep -Seconds 1
            $elapsed++
        }
        LogWarn "페이지 로드 대기 시간 초과 ($MaxWaitSec 초)"
        return $false
    }

    function Wait-Element {
        param([string]$wsUrl, [string]$Selector, [int]$MaxWaitSec=20)
        $expr = "!!document.querySelector('$Selector')"
        $elapsed = 0
        while ($elapsed -lt $MaxWaitSec) {
            try {
                $resp = Invoke-CDPEval -wsUrl $wsUrl -Expression $expr
                if ($resp.result.result.value -eq $true) { return $true }
            } catch {}
            Start-Sleep -Seconds 1
            $elapsed++
        }
        return $false
    }

    # ── 활성 탭 WebSocket URL 획득 ──────────────────────
    LogStep "3. 홈택스 자동 로그인"

    try {
        $tabsResp = Invoke-WebRequest -Uri "http://localhost:$DebugPort/json" -UseBasicParsing -TimeoutSec 5
        $tabs = $tabsResp.Content | ConvertFrom-Json
        $tab = $tabs | Where-Object { $_.type -eq "page" } | Select-Object -First 1
        if (-not $tab) { throw "활성 탭 없음" }
        $wsUrl = $tab.webSocketDebuggerUrl
        LogOK "탭 연결: $($tab.title)"
    } catch {
        LogErr "CDP 탭 목록 조회 실패: $_"
        exit 1
    }

    # ── 홈택스 공동인증서 로그인 페이지 이동 ────────────
    $loginPageUrl = "https://www.hometax.go.kr/websquareServlet?mv=UTEAHTSSPAB001M0"
    Log "홈택스 로그인 페이지 이동 중..."

    $navMsg = @{
        id     = 2
        method = "Page.navigate"
        params = @{ url = $loginPageUrl }
    }
    try {
        Invoke-CDP -wsUrl $wsUrl -Message $navMsg -TimeoutSec 30 | Out-Null
        Start-Sleep -Seconds 4
        Wait-PageReady -wsUrl $wsUrl -MaxWaitSec 30 | Out-Null
        LogOK "홈택스 로그인 페이지 로드"
    } catch {
        LogErr "홈택스 페이지 이동 실패: $_"
        Log "네트워크 연결 및 홈택스 접속 가능 여부를 확인하세요." Yellow
        exit 1
    }

    # ── 공동인증서 탭 클릭 ──────────────────────────────
    # 홈택스 로그인 화면에서 공동인증서 로그인 탭/버튼 클릭
    $clickCertTab = @"
(function(){
  // iframe 포함 전체 문서 탐색 헬퍼
  function allDocs() {
    var docs = [document];
    var frames = document.querySelectorAll('iframe');
    for(var i=0; i<frames.length; i++){
      try{ docs.push(frames[i].contentDocument || frames[i].contentWindow.document); }catch(e){}
    }
    return docs;
  }
  var selectors = [
    '#mf_txppWframe_login_tab7',
    '.w2troup.privLofin',
    '.login_tabcont.privLofin',
    'a[href*="cert"]',
    'li[id*="cert"]',
    '.login-tab:nth-child(2)',
    '#certLogin',
    'a.cert-login',
    '[title="공동인증서"]',
    '[alt="공동인증서"]'
  ];
  var docs = allDocs();
  for(var d=0; d<docs.length; d++){
    for(var i=0; i<selectors.length; i++){
      var el = docs[d].querySelector(selectors[i]);
      if(el){ el.click(); return 'clicked:'+selectors[i]+'(doc'+d+')'; }
    }
    // 텍스트로 탐색
    var links = docs[d].querySelectorAll('a, button, li');
    for(var j=0; j<links.length; j++){
      if(links[j].textContent.trim().indexOf('공동인증서') >= 0){
        links[j].click(); return 'clicked-text(doc'+d+'):'+links[j].textContent.trim().substring(0,30);
      }
    }
  }
  return 'not-found';
})()
"@
    Start-Sleep -Seconds 2
    $certTabResp = Invoke-CDPEval -wsUrl $wsUrl -Expression $clickCertTab -TimeoutSec 10
    Log "공동인증서 탭: $($certTabResp.result.result.value)"
    Start-Sleep -Seconds 2

    # ── 인택스 인증서 선택 ──────────────────────────────
    $certName = if ($IntaxCert) {
        # 인증서 디렉토리명에서 CN 추출 (cn= 이후, 첫 번째 쉼표 전)
        if ($IntaxCert.Name -match "cn=([^,]+)") { $Matches[1] } else { "인택스" }
    } else { "인택스" }

    $selectCertJs = @"
(function(){
  function allDocs() {
    var docs = [document];
    var frames = document.querySelectorAll('iframe');
    for(var i=0; i<frames.length; i++){
      try{ docs.push(frames[i].contentDocument || frames[i].contentWindow.document); }catch(e){}
    }
    return docs;
  }
  var docs = allDocs();
  for(var d=0; d<docs.length; d++){
    var certList = docs[d].querySelectorAll('.cert-list li, .certList li, [id*="certList"] li, #certList li, tr, .cert-row, .certItem');
    for(var i=0; i<certList.length; i++){
      if(certList[i].textContent.indexOf('인택스') >= 0 || certList[i].textContent.indexOf('97769') >= 0){
        certList[i].click();
        return 'cert-selected(doc'+d+'):'+certList[i].textContent.trim().substring(0,40);
      }
    }
  }
  return 'cert-not-found';
})()
"@

    Start-Sleep -Seconds 1
    $certSelResp = Invoke-CDPEval -wsUrl $wsUrl -Expression $selectCertJs -TimeoutSec 10
    Log "인증서 선택: $($certSelResp.result.result.value)"

    if ($certSelResp.result.result.value -eq "cert-not-found") {
        LogWarn "인택스 인증서를 자동으로 찾지 못했습니다."
        LogWarn "자동 선택 실패 — 홈택스 인증서 창에서 인택스(97769) 인증서를 수동으로 선택 후 스크립트를 재실행하세요."
        LogWarn "재실행 시 -SkipBrowser 없이 실행하면 로그인까지 자동으로 진행됩니다."
        exit 1
    }

    # ── PIN 입력 ─────────────────────────────────────────
    # 보안: PIN 값은 JS 문자열로 직접 삽입 (CDP 채널은 로컬 전용)
    $pinJs = @"
(function(pin){
  function allDocs() {
    var docs = [document];
    var frames = document.querySelectorAll('iframe');
    for(var i=0; i<frames.length; i++){
      try{ docs.push(frames[i].contentDocument || frames[i].contentWindow.document); }catch(e){}
    }
    return docs;
  }
  var pwSelectors = ['#certPw','#certPassword','#pin',
    'input[type="password"][id*="pin"]','input[type="password"][id*="cert"]',
    'input[type="password"]'];
  var docs = allDocs();
  for(var d=0; d<docs.length; d++){
    for(var s=0; s<pwSelectors.length; s++){
      var f = docs[d].querySelector(pwSelectors[s]);
      if(f){
        f.value = pin;
        f.dispatchEvent(new Event('input', {bubbles:true}));
        f.dispatchEvent(new Event('change', {bubbles:true}));
        return 'pin-entered(doc'+d+'):'+f.id;
      }
    }
  }
  return 'pin-field-not-found';
})('$CERT_PIN')
"@

    Start-Sleep -Seconds 1
    $pinResp = Invoke-CDPEval -wsUrl $wsUrl -Expression $pinJs -TimeoutSec 10
    Log "PIN 입력: $($pinResp.result.result.value)"

    # PIN 값 메모리에서 해제
    $CERT_PIN = $null
    [System.GC]::Collect()

    if ($pinResp.result.result.value -eq "pin-field-not-found") {
        LogErr "PIN 입력 필드를 찾지 못했습니다. 홈택스 페이지 구조 변경 가능성이 있습니다."
        Log "수동으로 PIN을 입력하고 로그인 후, 스크립트를 -SkipBrowser 옵션으로 다시 실행하세요." Yellow
        exit 1
    }

    # ── 확인(로그인) 버튼 클릭 ──────────────────────────
    $loginClickJs = @"
(function(){
  function allDocs() {
    var docs = [document];
    var frames = document.querySelectorAll('iframe');
    for(var i=0; i<frames.length; i++){
      try{ docs.push(frames[i].contentDocument || frames[i].contentWindow.document); }catch(e){}
    }
    return docs;
  }
  var btnSelectors = ['#logBtn','button[id*="login"]','a[id*="login"]',
    '.login-btn','[title="확인"]','input[type="submit"]'];
  var docs = allDocs();
  for(var d=0; d<docs.length; d++){
    for(var s=0; s<btnSelectors.length; s++){
      var b = docs[d].querySelector(btnSelectors[s]);
      if(b){ b.click(); return 'login-clicked(doc'+d+'):'+b.id; }
    }
    var all = docs[d].querySelectorAll('button, a, input[type="button"]');
    for(var j=0; j<all.length; j++){
      var t = all[j].textContent.trim();
      if(t==='확인' || t==='로그인'){
        all[j].click(); return 'text-clicked(doc'+d+'):'+t;
      }
    }
  }
  return 'login-btn-not-found';
})()
"@

    Start-Sleep -Seconds 1
    $loginResp = Invoke-CDPEval -wsUrl $wsUrl -Expression $loginClickJs -TimeoutSec 10
    Log "로그인 버튼: $($loginResp.result.result.value)"

    # 로그인 완료 대기 (최대 30초)
    Log "로그인 처리 중 (최대 30초 대기)..."
    Start-Sleep -Seconds 5
    Wait-PageReady -wsUrl $wsUrl -MaxWaitSec 25 | Out-Null

    # 로그인 성공 여부 확인
    $loginCheckJs = @"
(function(){
  var url = window.location.href;
  var fail = document.querySelector('.error-msg, .alert, [class*="error"]');
  if(fail && fail.textContent.trim().length > 0) return 'fail:'+fail.textContent.trim().substring(0,50);
  if(url.indexOf('hometax.go.kr') < 0) return 'navigated-away';
  return 'ok:'+url.substring(0,80);
})()
"@
    $loginCheck = Invoke-CDPEval -wsUrl $wsUrl -Expression $loginCheckJs -TimeoutSec 10
    $loginStatus = $loginCheck.result.result.value
    Log "로그인 상태: $loginStatus"

    if ($loginStatus -like "fail:*") {
        LogErr "로그인 실패: $loginStatus"
        Log "홈택스 비밀번호 또는 인증서 PIN을 확인하세요." Yellow
        LogErr "자동화 중단. 수동 로그인 후 -SkipBrowser 옵션으로 재실행하세요."
        exit 1
    }

    LogOK "홈택스 로그인 완료"

    # ── 수임거래처 현황 페이지 이동 ─────────────────────
    LogStep "4. 수임거래처 현황 페이지 이동"

    $clientListUrl = "https://www.hometax.go.kr/websquareServlet?mv=UTBFATBAA13Pp0"
    $navMsg2 = @{
        id     = 3
        method = "Page.navigate"
        params = @{ url = $clientListUrl }
    }
    try {
        Invoke-CDP -wsUrl $wsUrl -Message $navMsg2 -TimeoutSec 30 | Out-Null
        Start-Sleep -Seconds 5
        Wait-PageReady -wsUrl $wsUrl -MaxWaitSec 30 | Out-Null
        LogOK "수임거래처 현황 페이지 로드"
    } catch {
        LogWarn "수임거래처 직접 URL 실패. 메뉴 탐색 시도..."
        # 메뉴 탐색 대안
        $menuNavJs = @"
(function(){
  var links = document.querySelectorAll('a, li, span');
  for(var i=0; i<links.length; i++){
    var t = links[i].textContent.trim();
    if(t.indexOf('수임거래처') >= 0 || t.indexOf('수임 거래처') >= 0){
      links[i].click();
      return 'menu-clicked:'+t;
    }
  }
  return 'menu-not-found';
})()
"@
        $menuResp = Invoke-CDPEval -wsUrl $wsUrl -Expression $menuNavJs -TimeoutSec 10
        Log "메뉴 클릭: $($menuResp.result.result.value)"
        Start-Sleep -Seconds 5
        Wait-PageReady -wsUrl $wsUrl -MaxWaitSec 30 | Out-Null
    }

    # ── 엑셀 다운로드 버튼 클릭 ────────────────────────
    LogStep "5. 엑셀 다운로드"

    $downloadJs = @"
(function(){
  var candidates = [
    document.getElementById('mf_txppWframe_trigger44'),
    document.querySelector('.w2trigger[id*="trigger44"]'),
    document.querySelector('.w2trigger[id*="excel"]'),
    document.getElementById('excelDownload'),
    document.querySelector('[id*="excel"]'),
    document.querySelector('[id*="Excel"]'),
    document.querySelector('button[title*="엑셀"]'),
    document.querySelector('a[title*="엑셀"]'),
    document.querySelector('[onclick*="excel"]'),
    document.querySelector('[onclick*="Excel"]')
  ];
  for(var i=0; i<candidates.length; i++){
    if(candidates[i]){ candidates[i].click(); return 'excel-clicked:'+candidates[i].id; }
  }
  // 텍스트 탐색
  var all = document.querySelectorAll('button, a, input[type="button"], .w2trigger');
  for(var j=0; j<all.length; j++){
    var t = all[j].textContent.trim();
    if(t.indexOf('엑셀') >= 0 || t.indexOf('Excel') >= 0 || t.indexOf('EXCEL') >= 0){
      all[j].click();
      return 'text-clicked:'+t;
    }
  }
  return 'excel-btn-not-found';
})()
"@

    $dlResp = Invoke-CDPEval -wsUrl $wsUrl -Expression $downloadJs -TimeoutSec 10
    Log "다운로드 버튼: $($dlResp.result.result.value)"

    if ($dlResp.result.result.value -eq "excel-btn-not-found") {
        LogWarn "엑셀 다운로드 버튼을 찾지 못했습니다. 재시도 중..."
        Start-Sleep -Seconds 3

        # 재시도
        $dlResp2 = Invoke-CDPEval -wsUrl $wsUrl -Expression $downloadJs -TimeoutSec 10
        Log "재시도 결과: $($dlResp2.result.result.value)"

        if ($dlResp2.result.result.value -eq "excel-btn-not-found") {
            LogErr "엑셀 다운로드 버튼을 찾지 못했습니다."
            Log "홈택스 수임거래처 현황 페이지에서 수동으로 엑셀을 다운로드한 후," Yellow
            Log "-SkipBrowser 또는 -InputFile 옵션으로 재실행하세요." Yellow
            exit 1
        }
    }

    # 다운로드 완료 대기 (Downloads 폴더에서 최신 xlsx 파일 탐색)
    Log "다운로드 완료 대기 중..."
    Start-Sleep -Seconds 3

    $dlFile   = $null
    $dlWaited = 0
    while ($dlWaited -lt 60) {
        $recent = Get-ChildItem -Path $DownloadDir -Filter "*.xlsx" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($recent) { $dlFile = $recent.FullName; break }

        # xls도 확인
        $recentXls = Get-ChildItem -Path $DownloadDir -Filter "*.xls" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($recentXls) { $dlFile = $recentXls.FullName; break }

        Start-Sleep -Seconds 2
        $dlWaited += 2
    }

    if (-not $dlFile) {
        LogErr "다운로드된 엑셀 파일을 찾지 못했습니다 (60초 초과)."
        Log "수동으로 Downloads 폴더의 엑셀 파일 경로를 확인 후:" Yellow
        Log "  .\hometax_auto_sync.ps1 -InputFile `"다운로드된파일.xlsx`"" Yellow
        exit 1
    }

    $InputFile = $dlFile
    LogOK "다운로드 완료: $InputFile"

    # 비밀번호 변수도 해제
    $HOMETAX_PW = $null
    [System.GC]::Collect()
}

# ─── STEP 3: Excel → JSON 변환 ──────────────────────
LogStep "3. 홈택스 엑셀 → JSON 변환"

if (-not (Test-Path $ImportScript)) {
    LogErr "hometax_import.ps1을 찾을 수 없습니다: $ImportScript"
    exit 1
}

if ($InputFile -eq "" -or -not (Test-Path $InputFile)) {
    LogErr "변환할 입력 파일이 없습니다. -InputFile 옵션으로 엑셀 파일 경로를 지정하세요."
    exit 1
}

try {
    Log "변환 중: $InputFile -> $HtxDataFile"
    & $ImportScript -InputFile $InputFile -OutputFile $HtxDataFile
    if ($LASTEXITCODE -ne 0) { throw "hometax_import.ps1 종료 코드: $LASTEXITCODE" }
    LogOK "hometax_data.json 생성 완료"
} catch {
    LogErr "JSON 변환 실패: $_"
    exit 1
}

# ─── STEP 4: intax_backup.json 병합 ─────────────────
LogStep "4. intax_backup.json 병합"

if (-not (Test-Path $HtxDataFile)) {
    LogErr "hometax_data.json이 없습니다."
    exit 1
}

# 홈택스 데이터 로드
$htxRaw  = [System.IO.File]::ReadAllText($HtxDataFile, [System.Text.Encoding]::UTF8)
$htxData = $htxRaw | ConvertFrom-Json

if (-not $htxData.clients -or $htxData.clients.Count -eq 0) {
    LogErr "hometax_data.json에 거래처 데이터가 없습니다."
    exit 1
}

Log "홈택스 거래처 수: $($htxData.clients.Count)건"

# 기존 백업 로드
if (Test-Path $BackupFile) {
    $backupRaw = [System.IO.File]::ReadAllText($BackupFile, [System.Text.Encoding]::UTF8)
    $backup    = $backupRaw | ConvertFrom-Json
} else {
    LogWarn "intax_backup.json이 없습니다. 새로 생성합니다."
    $backup = [PSCustomObject]@{
        clients = @()
        staff   = @()
        users   = @()
        notices = @()
    }
}

# clients 배열을 해시테이블로 변환 (biz_no 기준 O(1) 조회)
$existingMap = @{}
if ($backup.clients) {
    foreach ($c in $backup.clients) {
        $key = ($c.biz_no -replace "[^0-9]", "")
        if ($key -ne "") {
            $existingMap[$key] = $c
        }
    }
}

$added   = 0
$updated = 0
$baseTs  = [System.DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
$idSeq   = 0

$divMap  = @{ "기장"="기장"; "신고대리"="신고대리"; "수임"="기장"; "신고"="신고대리" }
$typeMap = @{ "법인"="법인"; "개인"="개인"; "개인사업자"="개인" }

foreach ($hc in $htxData.clients) {
    $rawBizNo = if ($hc.biz_no) { $hc.biz_no } else { "" }
    $bizNo    = ($rawBizNo -replace "[^0-9\-]", "").Trim()
    $bizKey   = ($bizNo -replace "[^0-9]", "")
    $name     = if ($hc.name) { $hc.name } else { "(이름없음)" }

    # div/type/status 정규화
    $divRaw  = if ($hc.div)    { $hc.div }    else { "기장" }
    $typeRaw = if ($hc.type)   { $hc.type }   else { "개인" }
    $statRaw = if ($hc.status) { $hc.status } else { "정상" }

    $div    = if ($divMap[$divRaw])   { $divMap[$divRaw] }   else { "기장" }
    $type   = if ($typeMap[$typeRaw]) { $typeMap[$typeRaw] } else { "개인" }
    $status = if ($statRaw -match "폐업") { "폐업" } else { "정상" }
    $sector = if ($hc.sector) { $hc.sector } else { "" }

    if ($bizKey -ne "" -and $existingMap.ContainsKey($bizKey)) {
        # 기존 거래처 업데이트 (fee, staff_id, note 등 보존)
        $existing = $existingMap[$bizKey]
        $existing.name       = $name
        $existing.div        = $div
        $existing.type       = $type
        $existing.sector     = $sector
        $existing.status     = $status
        $existing.htx_synced = $true
        if ($hc.hometax_id) { $existing.hometax_id = $hc.hometax_id }
        $updated++
    } else {
        # 신규 거래처 추가
        $newId  = "c${baseTs}_${idSeq}"
        $idSeq++
        $newObj = [PSCustomObject]@{
            id          = $newId
            name        = $name
            div         = $div
            biz_no      = $bizNo
            type        = $type
            status      = $status
            sector      = $sector
            size        = if ($hc.size) { $hc.size } elseif ($type -eq "법인") { "중소기업" } else { "개인사업자" }
            fee         = 0
            cms         = "N"
            cms_status  = "미납"
            staff_id    = ""
            phone       = if ($hc.phone) { $hc.phone } else { "" }
            hometax_id  = if ($hc.hometax_id) { $hc.hometax_id } else { "" }
            karinara_id = ""
            note        = ""
            htx_synced  = $true
        }
        # clients 배열에 추가
        $backup.clients += $newObj
        if ($bizKey -ne "") { $existingMap[$bizKey] = $newObj }
        $added++
    }
}

$now = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")

# meta 섹션 업데이트 (없으면 추가)
$totalClients = $backup.clients.Count

# PSCustomObject에 meta 필드 추가 (Add-Member로)
if (-not (Get-Member -InputObject $backup -Name "meta" -MemberType NoteProperty -ErrorAction SilentlyContinue)) {
    $backup | Add-Member -MemberType NoteProperty -Name "meta" -Value ([PSCustomObject]@{
        version           = "2.0"
        exported_at       = $now
        hometax_synced_at = $now
        client_count      = $totalClients
    })
} else {
    $backup.meta.exported_at       = $now
    $backup.meta.hometax_synced_at = $now
    $backup.meta.client_count      = $totalClients
    if (-not $backup.meta.version) { $backup.meta | Add-Member -MemberType NoteProperty -Name "version" -Value "2.0" -Force }
}

# JSON으로 직렬화하여 저장 (Depth 10으로 중첩 구조 보존)
$outJson = $backup | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($BackupFile, $outJson, [System.Text.Encoding]::UTF8)

Log "신규 추가: ${added}건, 업데이트: ${updated}건" Green
LogOK "intax_backup.json 저장 완료 (총 $totalClients 건)"

# ─── STEP 5: Git Push ───────────────────────────────
if ($NoPush) {
    LogStep "5. Git push 건너뜀 (-NoPush 옵션)"
} else {
    LogStep "5. GitHub Push"

    $dateStr = (Get-Date).ToString("yyyy-MM-dd HH:mm")

    try {
        Push-Location $Root

        # git add (xlsx/xls는 .gitignore로 자동 제외)
        git add intax_backup.json hometax_data.json
        if ($LASTEXITCODE -ne 0) { throw "git add 실패" }

        # 변경사항 있는지 확인
        $status = git status --porcelain
        if (-not $status) {
            LogWarn "변경사항이 없습니다. 커밋 건너뜀."
        } else {
            git commit -m "data: 홈택스 자동 동기화 $dateStr (신규 $added, 업데이트 $updated)"
            if ($LASTEXITCODE -ne 0) { throw "git commit 실패" }

            git push
            if ($LASTEXITCODE -ne 0) { throw "git push 실패" }

            LogOK "GitHub 푸시 완료"
        }
    } catch {
        LogErr "Git 오류: $_"
        Log "수동으로 git push를 실행하세요:" Yellow
        Log "  git add intax_backup.json hometax_data.json" Yellow
        Log "  git commit -m `"data: 홈택스 수동 동기화 $dateStr`"" Yellow
        Log "  git push" Yellow
    } finally {
        Pop-Location
    }
}

# ─── 완료 ───────────────────────────────────────────
Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  홈택스 자동 동기화 완료" -ForegroundColor Cyan
Write-Host "  - 신규 거래처: ${added}건" -ForegroundColor Green
Write-Host "  - 업데이트:   ${updated}건" -ForegroundColor Green
Write-Host "  - 전체 거래처: ${totalClients}건" -ForegroundColor White
Write-Host "  - 동기화 시각: $now" -ForegroundColor White
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "대시보드 GitHub Pages가 자동으로 업데이트됩니다." -ForegroundColor Yellow
Write-Host "(GitHub Actions 빌드가 완료되면 반영됩니다)" -ForegroundColor Yellow
Write-Host ""
