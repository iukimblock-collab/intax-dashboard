# vault.ps1 - INTAX Credential Vault (Windows DPAPI)
# Encrypted data is only decryptable by the current Windows user on this PC.
#
# Usage:
#   .\vault.ps1 -Service hometax init       Encrypt .env(hometax).txt -> .hometax.vault
#   .\vault.ps1 -Service hometax show       Display stored credentials
#   .\vault.ps1 -Service hometax get PASSWORD   Get specific key value
#   .\vault.ps1 list                         List all vaults
#
# [홈택스 자동화 연동 설정]
# hometax_auto_sync.ps1을 사용하려면 .env(hometax).txt에 두 키를 모두 작성하세요:
#
#   PASSWORD=홈택스비밀번호
#   CERT_PIN=공동인증서PIN번호
#
# 작성 후 아래 명령으로 vault를 재생성하세요:
#   .\vault.ps1 -Service hometax init
#
# CERT_PIN은 DPAPI로 암호화되어 .hometax.vault에 저장됩니다.
# 절대 평문으로 스크립트나 로그에 기록하지 마세요.

param(
    [Parameter(Position=0)]
    [string]$Command = "list",

    [string]$Service = "",

    [Parameter(Position=1)]
    [string]$Key = ""
)

$Root = $PSScriptRoot

function Get-Paths([string]$service) {
    return @{
        Source = Join-Path $Root ".env($service).txt"
        Vault  = Join-Path $Root ".$service.vault"
    }
}

function Encrypt-String([string]$plainText) {
    $secure = ConvertTo-SecureString -String $plainText -AsPlainText -Force
    return ConvertFrom-SecureString $secure
}

function Decrypt-String([string]$encryptedText) {
    $secure = $encryptedText | ConvertTo-SecureString
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Load-Vault([string]$vaultPath) {
    if (-not (Test-Path $vaultPath)) {
        Write-Error "Vault not found: $vaultPath`nRun: .\vault.ps1 -Service $Service init"
        exit 1
    }
    $result = [ordered]@{}
    Get-Content $vaultPath | ForEach-Object {
        if ($_ -match "^([^#=]+)=(.+)$") {
            $result[$Matches[1].Trim()] = Decrypt-String $Matches[2].Trim()
        }
    }
    return $result
}

function Require-Service {
    if (-not $Service) {
        Write-Error "Specify -Service. Example: .\vault.ps1 -Service hometax show"
        exit 1
    }
}

switch ($Command.ToLower()) {

    "init" {
        Require-Service
        $paths = Get-Paths $Service

        if (-not (Test-Path $paths.Source)) {
            Write-Error "Source file not found: $($paths.Source)"
            exit 1
        }

        $lines = Get-Content $paths.Source
        $output = @(
            "# INTAX $Service credentials (Windows DPAPI encrypted)",
            "# Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm')",
            ""
        )

        foreach ($line in $lines) {
            if ($line -match "^([^#=]+)=(.+)$") {
                $k = $Matches[1].Trim()
                $v = $Matches[2].Trim()
                $output += "$k=$(Encrypt-String $v)"
                Write-Host "  Encrypted: $k" -ForegroundColor Green
            }
        }

        $output | Set-Content -Path $paths.Vault -Encoding UTF8
        Write-Host ""
        Write-Host "[Done] $($paths.Vault)" -ForegroundColor Cyan
    }

    "show" {
        Require-Service
        $paths = Get-Paths $Service
        $creds = Load-Vault $paths.Vault

        Write-Host ""
        Write-Host "=== $Service credentials ===" -ForegroundColor Cyan
        foreach ($k in $creds.Keys) {
            Write-Host "  $k = $($creds[$k])"
        }
        Write-Host ""
    }

    "get" {
        Require-Service
        if (-not $Key) {
            Write-Error "Specify a key name. Example: .\vault.ps1 -Service hometax get PASSWORD"
            exit 1
        }
        $paths = Get-Paths $Service
        $creds = Load-Vault $paths.Vault
        if ($creds.Contains($Key)) {
            Write-Output $creds[$Key]
        } else {
            Write-Error "Key not found: $Key  (available: $($creds.Keys -join ', '))"
            exit 1
        }
    }

    "list" {
        $vaults = Get-ChildItem -Path $Root -Filter "*.vault" -ErrorAction SilentlyContinue
        if ($vaults.Count -eq 0) {
            Write-Host "No vaults found." -ForegroundColor Yellow
        } else {
            Write-Host ""
            Write-Host "=== Stored vaults ===" -ForegroundColor Cyan
            foreach ($v in $vaults) {
                $svc = $v.BaseName.TrimStart(".")
                Write-Host "  $svc  ->  $($v.Name)" -ForegroundColor White
            }
            Write-Host ""
        }
    }

    default {
        Write-Host "Usage:"
        Write-Host "  .\vault.ps1 -Service [name] init         Encrypt source file"
        Write-Host "  .\vault.ps1 -Service [name] show         Display credentials"
        Write-Host "  .\vault.ps1 -Service [name] get [key]    Get specific key value"
        Write-Host "  .\vault.ps1 list                          List all vaults"
    }
}
