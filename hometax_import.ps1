# ================================================================
# hometax_import.ps1
# 홈택스 엑셀 다운로드 → hometax_data.json 변환 스크립트
#
# [사용법]
#   1. 홈택스(hometax.go.kr) 로그인
#   2. [세무대리인] → [수임동의 현황] → [수임거래처 조회] → 엑셀 다운로드
#   3. 이 스크립트 실행:
#        .\hometax_import.ps1 -InputFile "수임거래처목록.xlsx" [-OutputFile hometax_data.json]
#   4. 생성된 hometax_data.json 을 대시보드에서 [홈택스 가져오기]로 업로드
#
# [참고] 홈택스 엑셀 칼럼 (일반적인 형식):
#   A: 번호
#   B: 사업자등록번호
#   C: 상호(법인명)
#   D: 대표자명
#   E: 사업장소재지
#   F: 업태
#   G: 종목
#   H: 수임구분  (기장 / 신고대리)
#   I: 법인구분  (법인 / 개인)
#   J: 상태      (정상 / 폐업)
# ================================================================

param(
    [string]$InputFile  = "",
    [string]$OutputFile = "hometax_data.json",
    [switch]$Help
)

if($Help -or $InputFile -eq "") {
    Write-Host @"
사용법: .\hometax_import.ps1 -InputFile <엑셀파일경로> [-OutputFile <출력JSON경로>]

예시:
  .\hometax_import.ps1 -InputFile "C:\Downloads\수임거래처목록.xlsx"
  .\hometax_import.ps1 -InputFile "수임거래처.xlsx" -OutputFile "hometax_data.json"

출력 파일을 INTAX 대시보드 [거래처 현황] → [홈택스 가져오기] 버튼으로 업로드하세요.
"@
    exit 0
}

# ── 파일 존재 확인 ─────────────────────────────
if(-not (Test-Path $InputFile)) {
    Write-Error "파일을 찾을 수 없습니다: $InputFile"
    exit 1
}

$ext = [System.IO.Path]::GetExtension($InputFile).ToLower()

# ── Excel 읽기 (COM 객체 사용, Excel 설치 필요) ──
function Read-ExcelCOM($path) {
    $absPath = (Resolve-Path $path).Path
    Write-Host "Excel 파일 읽는 중: $absPath"

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($absPath)
        $ws = $wb.Worksheets.Item(1)

        $rows = @()
        $lastRow = $ws.UsedRange.Rows.Count

        # 헤더 행 찾기 (사업자등록번호 컬럼이 있는 행)
        $headerRow = 1
        for($r = 1; $r -le [Math]::Min(5, $lastRow); $r++) {
            $rowText = ""
            for($c = 1; $c -le 15; $c++) {
                $rowText += $ws.Cells.Item($r, $c).Text + "|"
            }
            if($rowText -match "사업자" -or $rowText -match "상호") {
                $headerRow = $r
                break
            }
        }

        # 헤더 매핑
        $headers = @{}
        for($c = 1; $c -le 20; $c++) {
            $h = $ws.Cells.Item($headerRow, $c).Text.Trim()
            if($h -ne "") { $headers[$h] = $c }
        }

        Write-Host "인식된 컬럼: $($headers.Keys -join ', ')"

        # 컬럼 인덱스 결정 (홈택스 표준 + 공통 변형 지원)
        # ※ ?? 는 PowerShell 7+ 전용이므로 5.1 호환 방식으로 작성
        if     ($headers["사업자등록번호"] -ne $null) { $colBiz  = $headers["사업자등록번호"] }
        elseif ($headers["사업자번호"]     -ne $null) { $colBiz  = $headers["사업자번호"]     }
        else                                          { $colBiz  = 2 }

        if     ($headers["상호(법인명)"]   -ne $null) { $colName = $headers["상호(법인명)"]   }
        elseif ($headers["상호"]           -ne $null) { $colName = $headers["상호"]           }
        elseif ($headers["법인명"]         -ne $null) { $colName = $headers["법인명"]         }
        else                                          { $colName = 3 }

        if     ($headers["종목"]           -ne $null) { $colSec  = $headers["종목"]           }
        elseif ($headers["업종"]           -ne $null) { $colSec  = $headers["업종"]           }
        else                                          { $colSec  = 7 }

        if     ($headers["업태"]           -ne $null) { $colBiz2 = $headers["업태"]           }  # 업태(sector fallback)
        else                                          { $colBiz2 = 6 }

        if     ($headers["수임구분"]       -ne $null) { $colDiv  = $headers["수임구분"]       }
        elseif ($headers["구분"]           -ne $null) { $colDiv  = $headers["구분"]           }
        else                                          { $colDiv  = 8 }

        if     ($headers["법인구분"]       -ne $null) { $colType = $headers["법인구분"]       }
        elseif ($headers["유형"]           -ne $null) { $colType = $headers["유형"]           }
        else                                          { $colType = 9 }

        if     ($headers["상태"]           -ne $null) { $colStat = $headers["상태"]           }
        else                                          { $colStat = 10 }

        $clients = @()
        for($r = $headerRow + 1; $r -le $lastRow; $r++) {
            $bizNo = $ws.Cells.Item($r, $colBiz).Text.Trim()
            $name  = $ws.Cells.Item($r, $colName).Text.Trim()
            if($bizNo -eq "" -and $name -eq "") { continue }

            $divRaw  = $ws.Cells.Item($r, $colDiv).Text.Trim()
            $typeRaw = $ws.Cells.Item($r, $colType).Text.Trim()
            $statRaw = $ws.Cells.Item($r, $colStat).Text.Trim()
            $secRaw  = $ws.Cells.Item($r, $colSec).Text.Trim()
            if($secRaw -eq "") { $secRaw = $ws.Cells.Item($r, $colBiz2).Text.Trim() }

            # 구분 정규화
            $div = switch -Wildcard ($divRaw) {
                "*기장*"    { "기장" }
                "*신고*"    { "신고대리" }
                "*수임*"    { "기장" }
                default     { "기장" }
            }
            # 유형 정규화
            $type = switch -Wildcard ($typeRaw) {
                "*법인*"    { "법인" }
                "*개인*"    { "개인" }
                default     { "개인" }
            }
            # 상태 정규화
            $status = if($statRaw -match "폐업") { "폐업" } else { "정상" }

            $clients += [PSCustomObject]@{
                biz_no     = $bizNo
                name       = $name
                type       = $type
                div        = $div
                sector     = $secRaw
                size       = if($type -eq "법인") { "중소기업" } else { "개인사업자" }
                status     = $status
                fee        = 0
                phone      = ""
                hometax_id = ""
                note       = "홈택스 가져오기"
            }
        }

        $wb.Close($false)
        return $clients
    }
    finally {
        if($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

# ── CSV 읽기 (Excel 없는 경우 대안) ──────────────
function Read-CSV($path) {
    Write-Host "CSV 파일 읽는 중: $path"
    $raw = Import-Csv -Path $path -Encoding UTF8
    $clients = @()
    foreach($row in $raw) {
        # ※ ?? 는 PowerShell 7+ 전용이므로 5.1 호환 방식으로 작성
        if     ($row."사업자등록번호" -ne $null) { $bizNo = $row."사업자등록번호".Trim() }
        elseif ($row."사업자번호"     -ne $null) { $bizNo = $row."사업자번호".Trim()     }
        else                                     { $bizNo = "" }

        if     ($row."상호(법인명)"   -ne $null) { $name = $row."상호(법인명)".Trim()   }
        elseif ($row."상호"           -ne $null) { $name = $row."상호".Trim()           }
        elseif ($row."법인명"         -ne $null) { $name = $row."법인명".Trim()         }
        else                                     { $name = "" }

        if($bizNo -eq "" -and $name -eq "") { continue }

        if     ($row."수임구분" -ne $null) { $divRaw  = $row."수임구분".Trim()  }
        elseif ($row."구분"     -ne $null) { $divRaw  = $row."구분".Trim()      }
        else                               { $divRaw  = "기장" }

        if     ($row."법인구분" -ne $null) { $typeRaw = $row."법인구분".Trim()  }
        elseif ($row."유형"     -ne $null) { $typeRaw = $row."유형".Trim()      }
        else                               { $typeRaw = "개인" }

        if     ($row."상태"     -ne $null) { $statRaw = $row."상태".Trim()      }
        else                               { $statRaw = "정상" }

        if     ($row."종목"     -ne $null) { $secRaw  = $row."종목".Trim()      }
        elseif ($row."업태"     -ne $null) { $secRaw  = $row."업태".Trim()      }
        else                               { $secRaw  = "" }
        $div = if($divRaw -match "신고") { "신고대리" } else { "기장" }
        $type = if($typeRaw -match "법인") { "법인" } else { "개인" }
        $status = if($statRaw -match "폐업") { "폐업" } else { "정상" }
        $clients += [PSCustomObject]@{
            biz_no = $bizNo; name = $name; type = $type; div = $div
            sector = $secRaw; size = if($type -eq "법인") {"중소기업"} else {"개인사업자"}
            status = $status; fee = 0; phone = ""; hometax_id = ""; note = "홈택스 가져오기"
        }
    }
    return $clients
}

# ── 실행 ─────────────────────────────────────────
Write-Host ""
Write-Host "=== 홈택스 가져오기 변환 스크립트 ===" -ForegroundColor Cyan
Write-Host ""

$clients = @()
if($ext -eq ".xlsx" -or $ext -eq ".xls") {
    try {
        $clients = Read-ExcelCOM $InputFile
    } catch {
        Write-Warning "Excel COM 방식 실패 (Excel 미설치?): $_"
        Write-Host "CSV 방식으로 재시도하려면 파일을 CSV로 저장 후 다시 실행하세요."
        exit 1
    }
} elseif($ext -eq ".csv") {
    $clients = Read-CSV $InputFile
} else {
    Write-Error "지원하지 않는 파일 형식: $ext (xlsx, xls, csv 만 지원)"
    exit 1
}

if($clients.Count -eq 0) {
    Write-Warning "변환된 거래처가 없습니다. 엑셀 파일 형식을 확인해 주세요."
    exit 1
}

$kijang    = ($clients | Where-Object { $_.div -eq "기장" }).Count
$sindae    = ($clients | Where-Object { $_.div -eq "신고대리" }).Count

Write-Host "변환 결과:" -ForegroundColor Green
Write-Host "  전체    : $($clients.Count)건"
Write-Host "  기장    : ${kijang}건"
Write-Host "  신고대리: ${sindae}건"
Write-Host ""

# JSON 출력
$output = [PSCustomObject]@{
    exported_at = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
    source      = "hometax"
    clients     = $clients
}

$json = $output | ConvertTo-Json -Depth 5
$absOutput = if ([System.IO.Path]::IsPathRooted($OutputFile)) { $OutputFile } else { Join-Path (Get-Location) $OutputFile }
[System.IO.File]::WriteAllText(
    $absOutput,
    $json,
    [System.Text.Encoding]::UTF8
)

Write-Host "저장 완료: $OutputFile" -ForegroundColor Green
Write-Host ""
Write-Host "다음 단계:" -ForegroundColor Yellow
Write-Host "  1. INTAX 대시보드 열기 (index.html)"
Write-Host "  2. [거래처 현황] 메뉴 이동"
Write-Host "  3. [홈택스 가져오기] 버튼 클릭"
Write-Host "  4. 생성된 '$OutputFile' 파일 선택"
Write-Host ""
exit 0
