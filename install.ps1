# PPT 편집기 설치: 이 폴더를 사용자 프로그램 폴더로 복사하고 바탕화면·시작 메뉴 바로가기를 만든다.
$ErrorActionPreference = 'Stop'
$src = $PSScriptRoot
$dst = Join-Path $env:LOCALAPPDATA 'Programs\PPT-Editor'

Write-Host ''
Write-Host '=== PPT 편집기 설치 ===' -ForegroundColor Cyan

if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application')) {
  Write-Host 'Microsoft PowerPoint(데스크톱 버전)가 설치되어 있지 않아요. PowerPoint를 먼저 설치해 주세요.' -ForegroundColor Red
  exit 1
}

# Stop a running copy of the editor so files can be replaced.
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -match [regex]::Escape($dst) } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }

New-Item -ItemType Directory -Force $dst | Out-Null
foreach ($item in @('app', 'engine', 'server.ps1', 'PPT편집기 실행.bat', '사용설명서.txt')) {
  $p = Join-Path $src $item
  if (Test-Path -LiteralPath $p) { Copy-Item -LiteralPath $p -Destination $dst -Recurse -Force }
}

$shell = New-Object -ComObject WScript.Shell
$targets = @(
  (Join-Path ([Environment]::GetFolderPath('Desktop')) 'PPT 편집기.lnk'),
  (Join-Path ([Environment]::GetFolderPath('Programs')) 'PPT 편집기.lnk')
)
foreach ($t in $targets) {
  $lnk = $shell.CreateShortcut($t)
  $lnk.TargetPath = Join-Path $dst 'PPT편집기 실행.bat'
  $lnk.WorkingDirectory = $dst
  $lnk.IconLocation = "$env:SystemRoot\System32\imageres.dll,170"
  $lnk.WindowStyle = 7
  $lnk.Description = 'PPT 편집기'
  $lnk.Save()
}

$fonts = (New-Object -ComObject Shell.Application).Namespace(0x14).Items() | ForEach-Object { $_.Name }
if (-not ($fonts -match 'Pretendard')) {
  Write-Host ''
  Write-Host '참고: Pretendard 글꼴이 없어요. 기본 글꼴(맑은 고딕)로 동작하며, 원하면 https://github.com/orioncactus/pretendard 에서 설치할 수 있어요.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host "설치 완료: $dst" -ForegroundColor Green
Write-Host '바탕화면의 "PPT 편집기" 아이콘으로 실행하세요. 만든 PPT는 [문서\PPT Editor] 폴더에 저장돼요.'
