<#
  solidworks-mcp 설치 / 업데이트

  install.cmd 를 더블클릭하거나 PowerShell에서:
      powershell -ExecutionPolicy Bypass -File install.ps1 [-Update] [-NoDesktop] [-NoClaudeCode] [-Python <python.exe>]

  하는 일: Python 3.11+ 찾기 -> .venv 만들기 -> 패키지 설치 -> Claude Desktop 설정에 등록
           -> (claude CLI가 있으면) Claude Code에 등록 -> 설치 점검(doctor.py)
  다시 실행해도 안전합니다. -Update 는 먼저 git pull 로 최신 코드를 받습니다.
#>
param(
    [switch]$Update,
    [switch]$NoDesktop,
    [switch]$NoClaudeCode,
    [string]$Python = ""
)

# Native tools (pip, git) write warnings to stderr; judge them by exit code only.
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Venv = Join-Path $Root ".venv"
$VenvPy = Join-Path $Venv "Scripts\python.exe"
$Server = Join-Path $Root "server.py"

function Step($text) { Write-Host "`n==> $text" -ForegroundColor Cyan }
function Fail($text) { Write-Host "`n[실패] $text" -ForegroundColor Red; exit 1 }

function Test-Python($exe) {
    # "3.12" 같은 버전 문자열, 3.11 미만이거나 실행이 안 되면 $null
    try {
        $v = & $exe -c "import sys; print('%d.%d' % sys.version_info[:2]) if sys.version_info >= (3, 11) else None" 2>$null
        if ($LASTEXITCODE -eq 0 -and $v -and $v -ne "None") { return $v.Trim() }
    } catch { }
    return $null
}

function Find-Python {
    if ($Python) {
        if (Test-Python $Python) { return $Python }
        Fail "지정한 Python($Python)이 3.11 이상이 아니거나 실행되지 않습니다."
    }
    # py 런처: 설치된 Python 중 3.11 이상의 가장 새 버전
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        $list = & py -0p 2>$null
        $cands = @()
        foreach ($line in $list) {
            if ($line -match '-V:(\d+)\.(\d+)\S*\s+\*?\s*(.+python\.exe)') {
                $cands += [pscustomobject]@{ Major = [int]$matches[1]; Minor = [int]$matches[2]; Path = $matches[3].Trim() }
            }
        }
        foreach ($c in ($cands | Sort-Object Major, Minor -Descending)) {
            if (($c.Major -gt 3 -or ($c.Major -eq 3 -and $c.Minor -ge 11)) -and (Test-Python $c.Path)) { return $c.Path }
        }
    }
    foreach ($name in @("python", "python3")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -notlike "*WindowsApps*" -and (Test-Python $cmd.Source)) { return $cmd.Source }
    }
    return $null
}

Write-Host "solidworks-mcp 설치 ($Root)" -ForegroundColor Green

if ($Update) {
    Step "최신 코드 받기 (git pull)"
    if ((Test-Path (Join-Path $Root ".git")) -and (Get-Command git -ErrorAction SilentlyContinue)) {
        & git -C $Root pull --ff-only
        if ($LASTEXITCODE -ne 0) { Fail "git pull 실패. 로컬에서 고친 파일이 있으면 정리한 뒤 다시 실행하세요." }
    } else {
        Write-Host "git 저장소가 아니어서 건너뜁니다. (ZIP으로 받았다면 새 ZIP을 같은 폴더에 덮어쓰세요.)"
    }
}

Step "Python 3.11 이상 찾기"
$Base = Find-Python
if (-not $Base) {
    Write-Host "Python 3.11 이상이 없습니다. 아래 중 하나로 설치한 뒤 다시 실행하세요:" -ForegroundColor Yellow
    Write-Host "  - https://www.python.org/downloads/ (설치 화면에서 'Add python.exe to PATH' 체크)"
    Write-Host "  - winget install Python.Python.3.12"
    exit 1
}
Write-Host "사용할 Python: $Base ($(Test-Python $Base))"

Step "가상환경 준비 (.venv)"
if ((Test-Path $VenvPy) -and -not (Test-Python $VenvPy)) {
    Write-Host "기존 .venv가 깨졌거나 너무 오래되어 다시 만듭니다."
    Remove-Item -Recurse -Force $Venv
}
if (-not (Test-Path $VenvPy)) {
    & $Base -m venv $Venv
    if ($LASTEXITCODE -ne 0) { Fail "가상환경을 만들지 못했습니다." }
}

Step "패키지 설치 (mcp, pywin32, pillow)"
& $VenvPy -m pip install --disable-pip-version-check --quiet --upgrade pip
& $VenvPy -m pip install --disable-pip-version-check --quiet -e $Root
if ($LASTEXITCODE -ne 0) { Fail "패키지 설치 실패. 인터넷 연결(회사 프록시)을 확인하세요." }

if (-not $NoDesktop) {
    Step "Claude Desktop 설정에 등록"
    & $VenvPy (Join-Path $Root "doctor.py") --register
}

if (-not $NoClaudeCode) {
    $claude = Get-Command claude -ErrorAction SilentlyContinue
    if ($claude) {
        Step "Claude Code에 등록"
        & claude mcp get solidworks *> $null
        if ($LASTEXITCODE -eq 0) { & claude mcp remove solidworks --scope user *> $null }
        & claude mcp add solidworks --scope user -- $VenvPy $Server
    }
}

Step "설치 점검"
& $VenvPy (Join-Path $Root "doctor.py")
$ok = ($LASTEXITCODE -eq 0)

Write-Host ""
if ($ok) {
    Write-Host "설치가 끝났습니다." -ForegroundColor Green
    Write-Host "Claude Desktop을 완전히 종료(트레이 아이콘 > 종료)한 뒤 다시 켜면 'solidworks' 도구가 보입니다."
} else {
    Write-Host "점검에서 [실패] 항목이 있습니다. 위 안내대로 고친 뒤 install.cmd를 다시 실행하세요." -ForegroundColor Yellow
}
