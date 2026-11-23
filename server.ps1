# 원클라우드 현황판 - 로컬 실행 프로그램
# 이 폴더의 index.html 을 내 PC 안(localhost)에서만 띄우고, 사용현황 엑셀을 읽어 전달합니다.
# 창을 닫으면 1분 안에 스스로 종료됩니다.

$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LibDir = Join-Path $AppDir 'lib'
$LogFile = Join-Path $AppDir 'server.log'
function Log($m) { try { Add-Content -Path $LogFile -Value ("{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date), $m) -Encoding UTF8 } catch {} }

# ---- 설정 ----
$cfg = [ordered]@{ excel = '..\원클라우드 계약서관리.xlsx'; payment = '..\8. 입금내역\입금내역정리.xlsx'; port = 8765; browser = 'edge-app' }
$cfgPath = Join-Path $AppDir 'config.json'
if (Test-Path $cfgPath) {
  try {
    $j = Get-Content -Path $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in @('excel','payment','port','browser')) { if ($j.$k) { $cfg[$k] = $j.$k } }
  } catch { Log "config.json 읽기 실패: $_" }
}
function Get-ExcelPath([string]$which = 'use') {
  $p = if ($which -eq 'pay') { $cfg.payment } else { $cfg.excel }
  if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $AppDir $p }
  return [System.IO.Path]::GetFullPath($p)
}
$Url = "http://localhost:$($cfg.port)/"

function Open-Dashboard {
  $opened = $false
  if ($cfg.browser -eq 'edge-app') {
    try { Start-Process -FilePath 'msedge' -ArgumentList "--app=$Url", '--window-size=1400,950'; $opened = $true } catch {}
    if (-not $opened) { try { Start-Process -FilePath 'chrome' -ArgumentList "--app=$Url", '--window-size=1400,950'; $opened = $true } catch {} }
  }
  if (-not $opened) { Start-Process $Url }
}

# ---- 바로가기(바탕화면 + 이 폴더) 만들기: 처음 한 번 ----
try {
  $ws = New-Object -ComObject WScript.Shell
  $icon = Join-Path (Split-Path $AppDir -Parent) 'oneCloud_1.ico'
  foreach ($dir in @([Environment]::GetFolderPath('Desktop'), $AppDir)) {
    $lnkPath = Join-Path $dir '원클라우드 현황판.lnk'
    if (-not (Test-Path $lnkPath)) {
      $lnk = $ws.CreateShortcut($lnkPath)
      $lnk.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
      $lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MyInvocation.MyCommand.Path)`""
      $lnk.WorkingDirectory = $AppDir
      $lnk.WindowStyle = 7
      if (Test-Path $icon) { $lnk.IconLocation = "$icon,0" }
      $lnk.Description = '원클라우드 사용 현황판'
      $lnk.Save()
    }
  }
} catch { Log "바로가기 생성 실패: $_" }

# ---- 서버 시작 (이미 켜져 있으면 창만 열기) ----
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Url)
try { $listener.Start() } catch {
  Log "이미 실행 중이거나 포트 사용 중: $_"
  Open-Dashboard
  exit
}
Log "시작: $Url  엑셀: $(Get-ExcelPath)"

# ---- 라이브러리 준비 (처음 한 번 인터넷에서 받아 lib 폴더에 보관 → 이후 오프라인 동작) ----
$Libs = @{
  'south-korea.js'   = 'https://cdn.jsdelivr.net/npm/@svg-maps/south-korea@1/index.js'
}
function Ensure-Lib($name) {
  $f = Join-Path $LibDir $name
  if (Test-Path $f) { return $f }
  if (-not $Libs.ContainsKey($name)) { return $null }
  try {
    if (-not (Test-Path $LibDir)) { New-Item -ItemType Directory -Path $LibDir | Out-Null }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $Libs[$name] -OutFile $f -UseBasicParsing -TimeoutSec 30
    Log "라이브러리 저장: $name"
    return $f
  } catch { Log "라이브러리 다운로드 실패($name): $_"; return $null }
}

$Mime = @{ '.html'='text/html; charset=utf-8'; '.js'='text/javascript; charset=utf-8'; '.json'='application/json; charset=utf-8'; '.ico'='image/x-icon'; '.css'='text/css; charset=utf-8'; '.png'='image/png' }

function Send($ctx, [int]$code, [byte[]]$bytes, [string]$type) {
  $r = $ctx.Response
  $r.StatusCode = $code
  $r.ContentType = $type
  $r.Headers.Add('Cache-Control', 'no-store')
  $r.ContentLength64 = $bytes.Length
  if ($bytes.Length) { $r.OutputStream.Write($bytes, 0, $bytes.Length) }
  $r.OutputStream.Close()
}
function SendText($ctx, [int]$code, [string]$text, [string]$type = 'text/plain; charset=utf-8') {
  Send $ctx $code ([Text.Encoding]::UTF8.GetBytes($text)) $type
}
function Read-Shared([string]$path) {
  # 엑셀에서 파일을 열어 둔 상태에서도 읽을 수 있도록 공유 모드로 읽기
  $fs = [System.IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
  try { $ms = New-Object System.IO.MemoryStream; $fs.CopyTo($ms); return $ms.ToArray() } finally { $fs.Close() }
}

Open-Dashboard

$IdleLimit = [TimeSpan]::FromSeconds(75)
$last = Get-Date
$pending = $listener.GetContextAsync()
while ($true) {
  if (-not $pending.Wait(1000)) {
    if (((Get-Date) - $last) -gt $IdleLimit) { Log '창이 닫혀 종료'; break }
    continue
  }
  $ctx = $pending.Result
  $pending = $listener.GetContextAsync()
  $last = Get-Date
  try {
    $path = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath)
    $which = $ctx.Request.QueryString['f']
    if ($which -ne 'pay') { $which = 'use' }
    switch -Regex ($path) {
      '^/api/stat$' {
        $x = Get-ExcelPath $which
        if (Test-Path -LiteralPath $x) {
          $fi = Get-Item -LiteralPath $x
          $ms = [long](($fi.LastWriteTimeUtc - [datetime]'1970-01-01').TotalMilliseconds)
          $o = @{ found = $true; name = $fi.Name; path = $fi.FullName; mtime = $ms; size = $fi.Length }
        } else { $o = @{ found = $false; path = $x } }
        SendText $ctx 200 ($o | ConvertTo-Json -Compress) 'application/json; charset=utf-8'
        break
      }
      '^/api/data$' {
        $x = Get-ExcelPath $which
        $bytes = $null
        for ($i = 0; $i -lt 5 -and -not $bytes; $i++) { try { $bytes = Read-Shared $x } catch { Start-Sleep -Milliseconds 400 } }
        if ($bytes) { Send $ctx 200 $bytes 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
        else { SendText $ctx 503 'busy' }
        break
      }
      '^/api/open$' {
        try { Start-Process -FilePath (Get-ExcelPath $which) } catch { Log "엑셀 열기 실패: $_" }
        SendText $ctx 200 'ok'
        break
      }
      '^/lib/(.+)$' {
        $name = [IO.Path]::GetFileName($Matches[1])
        $f = Ensure-Lib $name
        if ($f) { Send $ctx 200 ([IO.File]::ReadAllBytes($f)) 'text/javascript; charset=utf-8' } else { SendText $ctx 404 'not found' }
        break
      }
      default {
        $rel = if ($path -eq '/' -or $path -eq '') { 'index.html' } else { $path.TrimStart('/') }
        $full = [IO.Path]::GetFullPath((Join-Path $AppDir $rel))
        if ($full.StartsWith($AppDir, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $full -PathType Leaf)) {
          $ext = [IO.Path]::GetExtension($full).ToLower()
          $type = if ($Mime.ContainsKey($ext)) { $Mime[$ext] } else { 'application/octet-stream' }
          Send $ctx 200 ([IO.File]::ReadAllBytes($full)) $type
        } else { SendText $ctx 404 'not found' }
      }
    }
  } catch {
    Log "요청 처리 오류: $_"
    try { SendText $ctx 500 'error' } catch {}
  }
}
$listener.Stop()
$listener.Close()
