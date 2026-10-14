# PPT Editor local server: serves the editor UI on http://localhost and builds decks with the bundled engine.
$ErrorActionPreference = 'Stop'
$AppDir    = $PSScriptRoot
$EngineDir = Join-Path $AppDir 'engine'
$DataDir   = Join-Path $env:APPDATA 'PPT-Editor'
$ProjDir   = Join-Path $DataDir 'projects'
$UpDir     = Join-Path $DataDir 'uploads'
$OutDir    = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PPT Editor'
$WorkRoot  = Join-Path $env:TEMP 'ppt-editor'
$BrandsPath  = Join-Path $DataDir 'brands.json'
$HistoryPath = Join-Path $DataDir 'history.json'
New-Item -ItemType Directory -Force $DataDir, $ProjDir, $UpDir, $OutDir, $WorkRoot | Out-Null
# One-time migration from the first version, which kept brands next to the app.
$oldBrands = Join-Path $AppDir 'brands.json'
if ((Test-Path $oldBrands) -and -not (Test-Path $BrandsPath)) { Move-Item $oldBrands $BrandsPath }
$utf8 = New-Object Text.UTF8Encoding $false
Add-Type -AssemblyName System.Drawing

function ReadText([string]$p) { return [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) }
function WriteText([string]$p, [string]$t) { [IO.File]::WriteAllText($p, $t, $utf8) }
function Send($ctx, [int]$code, [string]$type, [byte[]]$bytes) {
  $r = $ctx.Response
  $r.StatusCode = $code; $r.ContentType = $type
  $r.Headers['Cache-Control'] = 'no-store'
  $r.ContentLength64 = $bytes.Length
  $r.OutputStream.Write($bytes, 0, $bytes.Length); $r.OutputStream.Close()
}
function SendJson($ctx, $obj, [int]$code = 200) { Send $ctx $code 'application/json; charset=utf-8' ($utf8.GetBytes((ConvertTo-Json -InputObject $obj -Depth 30 -Compress))) }
function SendRaw($ctx, [string]$json) { Send $ctx 200 'application/json; charset=utf-8' ($utf8.GetBytes($json)) }
$MIME = @{ '.html'='text/html; charset=utf-8'; '.js'='text/javascript; charset=utf-8'; '.css'='text/css; charset=utf-8'; '.png'='image/png'; '.jpg'='image/jpeg'; '.jpeg'='image/jpeg'; '.gif'='image/gif'; '.bmp'='image/bmp'; '.svg'='image/svg+xml'; '.json'='application/json; charset=utf-8' }
function SendFile($ctx, [string]$path) {
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Send $ctx 404 'text/plain' ($utf8.GetBytes('not found')); return }
  $type = $MIME[[IO.Path]::GetExtension($path).ToLower()]; if (-not $type) { $type = 'application/octet-stream' }
  Send $ctx 200 $type ([IO.File]::ReadAllBytes($path))
}
function BodyText($ctx) { $sr = New-Object IO.StreamReader($ctx.Request.InputStream, [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $sr.Close(); return $t }
function BodyBytes($ctx) { $ms = New-Object IO.MemoryStream; $ctx.Request.InputStream.CopyTo($ms); return $ms.ToArray() }
function SafeName([string]$n) { $n = ($n -replace '[\\/:*?"<>|\r\n]', '_').Trim().TrimEnd('.'); if (-not $n) { $n = 'presentation' }; return $n }
function RawOr([string]$path, [string]$fallback) { if (Test-Path -LiteralPath $path) { return (ReadText $path) } else { return $fallback } }
function IsUnder([string]$path, [string]$root) { $full = [IO.Path]::GetFullPath($path); return $full.StartsWith([IO.Path]::GetFullPath($root).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) }
function JStr([string]$s) { return (ConvertTo-Json -InputObject $s) }
# Query values decoded as UTF-8 (HttpListener's QueryString garbles non-ASCII names such as Korean file names).
function QS($ctx, [string]$key) {
  foreach ($pair in $ctx.Request.Url.Query.TrimStart('?').Split('&')) { $kv = $pair.Split('=', 2); if ($kv.Count -eq 2 -and [Uri]::UnescapeDataString($kv[0]) -eq $key) { return [Uri]::UnescapeDataString($kv[1].Replace('+', ' ')) } }
  return ''
}

$FontsJson = ConvertTo-Json -InputObject @((New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name } | Where-Object { -not $_.StartsWith('@') } | Sort-Object -Unique) -Compress
# Localized family names (e.g. '페이퍼로지 6 SemiBold') -> the name browsers know ('Paperlogy 6 SemiBold').
# Files saved on Korean Office store the Korean name, which the browser cannot resolve. Built on first use (a few seconds).
$script:FontAliasJson = $null
function FontAlias {
  if ($script:FontAliasJson) { return $script:FontAliasJson }
  Add-Type -AssemblyName PresentationCore
  $map = @{}; $all = New-Object 'System.Collections.Generic.HashSet[string]'
  $dirs = @("$env:windir\Fonts", "$env:LOCALAPPDATA\Microsoft\Windows\Fonts")
  foreach ($f in (Get-ChildItem $dirs -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(ttf|otf|ttc)$' })) {
    try {
      $g = New-Object Windows.Media.GlyphTypeface ([Uri]$f.FullName)
      $names = $g.Win32FamilyNames; $en = $null
      foreach ($kv in $names.GetEnumerator()) { if ($kv.Key.Name -eq 'en-US') { $en = $kv.Value } }
      if (-not $en) { continue }
      foreach ($kv in $names.GetEnumerator()) { [void]$all.Add($kv.Value); if ($kv.Value -ne $en -and -not $map.ContainsKey($kv.Value)) { $map[$kv.Value] = $en } }
    } catch {}
  }
  $aliasJson = if ($map.Count) { ConvertTo-Json -InputObject $map -Compress } else { '{}' }
  # names: every installed family name in any language, to tell which fonts of an imported file are missing
  $script:FontAliasJson = '{"alias":' + $aliasJson + ',"names":' + (ConvertTo-Json -InputObject @($all) -Compress) + '}'
  return $script:FontAliasJson
}

# Make an all-white copy of a transparent logo for dark backgrounds. Returns $null if the image has no transparency.
function WhiteVersion([string]$src, [string]$dst) {
  $bmp = New-Object System.Drawing.Bitmap $src
  try {
    $hasAlpha = $false
    $out = New-Object System.Drawing.Bitmap $bmp.Width, $bmp.Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($out); $g.DrawImage($bmp, 0, 0, $bmp.Width, $bmp.Height); $g.Dispose()
    $rect = New-Object System.Drawing.Rectangle 0, 0, $out.Width, $out.Height
    $data = $out.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadWrite, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $len = $data.Stride * $out.Height
    $buf = New-Object byte[] $len
    [Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $len)
    for ($i = 0; $i -lt $len; $i += 4) { if ($buf[$i+3] -lt 250) { $hasAlpha = $true }; $buf[$i] = 255; $buf[$i+1] = 255; $buf[$i+2] = 255 }
    [Runtime.InteropServices.Marshal]::Copy($buf, 0, $data.Scan0, $len)
    $out.UnlockBits($data)
    if ($hasAlpha) { $out.Save($dst, [System.Drawing.Imaging.ImageFormat]::Png) }
    $out.Dispose()
    if ($hasAlpha) { return $dst } else { return $null }
  } finally { $bmp.Dispose() }
}
function ImageAspect([string]$path) {
  if ($path.ToLower().EndsWith('.svg')) {
    $t = ReadText $path
    if ($t -match 'viewBox="\s*[\d\.\-]+[\s,]+[\d\.\-]+[\s,]+([\d\.]+)[\s,]+([\d\.]+)') { return [double]$Matches[1] / [double]$Matches[2] }
    if ($t -match 'width="([\d\.]+)[a-z]*"[^>]*height="([\d\.]+)') { return [double]$Matches[1] / [double]$Matches[2] }
    return 4
  }
  $img = [System.Drawing.Image]::FromFile($path); try { return $img.Width / [double]$img.Height } finally { $img.Dispose() }
}

function Upload($ctx) {
  $q = $ctx.Request.QueryString
  $orig = QS $ctx 'name'; $kind = QS $ctx 'kind'
  $ext = [IO.Path]::GetExtension($orig).ToLower()
  $allowed = @('.png','.jpg','.jpeg','.gif','.bmp'); if ($kind -eq 'logo') { $allowed += '.svg' }
  if ($allowed -notcontains $ext) { SendJson $ctx @{ error = "지원하지 않는 형식: $ext" } 400; return }
  $bytes = BodyBytes $ctx
  if ($bytes.Length -gt 30MB) { SendJson $ctx @{ error = 'file too large (30MB max)' } 400; return }
  $base = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,6)
  $file = Join-Path $UpDir ($base + $ext)
  [IO.File]::WriteAllBytes($file, $bytes)
  try { $aspect = ImageAspect $file } catch { Remove-Item -LiteralPath $file; SendJson $ctx @{ error = '이미지를 읽을 수 없어요.' } 400; return }
  $res = [ordered]@{ path = $file; url = "/uploads/$base$ext"; aspect = [math]::Round($aspect, 4); name = [IO.Path]::GetFileName($orig) }
  if ($kind -eq 'logo' -and $ext -ne '.svg' -and $ext -ne '.jpg' -and $ext -ne '.jpeg') {
    try { $w = WhiteVersion $file (Join-Path $UpDir ($base + '-white.png')); if ($w) { $res.white = $w; $res.whiteUrl = "/uploads/$base-white.png" } } catch {}
  }
  SendJson $ctx $res
}

function Build($ctx) {
  $req = (BodyText $ctx) | ConvertFrom-Json
  $name = SafeName ([string]$req.fileName)
  $jobName = (Get-Date -Format 'yyyyMMdd-HHmmss')
  $job = Join-Path $WorkRoot $jobName
  New-Item -ItemType Directory -Force $job | Out-Null
  WriteText (Join-Path $job 'spec.json')  ([string]$req.specJson)
  WriteText (Join-Path $job 'style.json') ([string]$req.styleJson)
  WriteText (Join-Path $job 'brand.json') ([string]$req.brandJson)
  $out = Join-Path $OutDir ($name + '.pptx')
  $i = 2
  while (Test-Path -LiteralPath $out) {
    try { $fs = [IO.File]::Open($out, 'Open', 'ReadWrite', 'None'); $fs.Close(); break } catch { $out = Join-Path $OutDir ("$name-$i.pptx"); $i++ }
  }
  $q = { param($p) '"' + $p + '"' }
  $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (& $q (Join-Path $EngineDir 'scripts\build.ps1')),
    '-Spec', (& $q (Join-Path $job 'spec.json')), '-Style', (& $q (Join-Path $job 'style.json')),
    '-Brand', (& $q (Join-Path $job 'brand.json')), '-Out', (& $q $out), '-PngDir', (& $q (Join-Path $job 'png')),
    '-HistoryPath', (& $q $HistoryPath))
  if ($req.potx) { $argList += '-Potx' }
  Write-Host ("[build] " + $name + " ...")
  $null = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput (Join-Path $job 'log.txt') -RedirectStandardError (Join-Path $job 'err.txt')
  $log = (RawOr (Join-Path $job 'log.txt') '') + (RawOr (Join-Path $job 'err.txt') '')
  $ok = ($log -match '(?m)^OK ')
  $pngs = @()
  if (Test-Path (Join-Path $job 'png')) { $pngs = @(Get-ChildItem (Join-Path $job 'png') -Filter 'slide-*.png' | Sort-Object Name | ForEach-Object { "/png/$jobName/" + $_.Name }) }
  Write-Host ("[build] " + $(if ($ok) { 'OK ' + $out } else { 'FAILED' }))
  SendJson $ctx ([ordered]@{ ok = [bool]$ok; out = $out; potx = $(if ($req.potx) { [IO.Path]::ChangeExtension($out, '.potx') } else { $null }); pngs = $pngs; log = $log.Trim() })
}

# ---------- Existing presentations: import and edit in place ----------
$ImpDir = Join-Path $DataDir 'imports'
New-Item -ItemType Directory -Force $ImpDir | Out-Null
function RunEngine([string]$script, [string[]]$scriptArgs, [string]$logBase) {
  $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + (Join-Path $EngineDir "scripts\$script") + '"')) + ($scriptArgs | ForEach-Object { if ($_ -like '-*') { $_ } else { '"' + $_ + '"' } })
  $null = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -NoNewWindow -Wait -PassThru -RedirectStandardOutput "$logBase.log" -RedirectStandardError "$logBase.err"
  return ((RawOr "$logBase.log" '') + (RawOr "$logBase.err" '')).Trim()
}
function UniqueOut([string]$name) {
  $out = Join-Path $OutDir ($name + '.pptx'); $i = 2
  while (Test-Path -LiteralPath $out) { $out = Join-Path $OutDir ("$name-$i.pptx"); $i++ }
  return $out
}
# Windows PowerShell's ConvertFrom-Json hands back a JSON array as one object, and ConvertTo-Json wraps nested
# arrays as {value, Count}; flatten both so history.json stays a plain list of entries.
function FlatHistory($node) {
  $out = @()
  foreach ($h in @($node)) {
    if ($null -eq $h) { continue }
    if ($h -is [array]) { $out += @(FlatHistory $h) }
    elseif ($h.PSObject.Properties['value'] -and $h.PSObject.Properties['Count'] -and -not $h.PSObject.Properties['date']) { $out += @(FlatHistory $h.value) }
    elseif ($h -is [string]) { continue }
    else { $out += $h }
  }
  return $out
}
function RenderDoc($ctx) {
  $req = (BodyText $ctx) | ConvertFrom-Json
  $name = SafeName ([string]$req.fileName)
  $jobName = (Get-Date -Format 'yyyyMMdd-HHmmss')
  $job = Join-Path $WorkRoot $jobName
  New-Item -ItemType Directory -Force $job | Out-Null
  WriteText (Join-Path $job 'doc.json') ([string]$req.docJson)
  # overwrite an earlier file with the same name unless it is open in PowerPoint
  $out = Join-Path $OutDir ($name + '.pptx'); $i = 2
  while (Test-Path -LiteralPath $out) { try { $fs = [IO.File]::Open($out, 'Open', 'ReadWrite', 'None'); $fs.Close(); break } catch { $out = Join-Path $OutDir ("$name-$i.pptx"); $i++ } }
  $args2 = @('-Doc', (Join-Path $job 'doc.json'), '-Out', $out, '-PngDir', (Join-Path $job 'png'))
  if ($req.potx) { $args2 += '-Potx' }
  if ($req.pdf) { $args2 += '-Pdf' }
  Write-Host ("[render] " + $name + " ...")
  $log = RunEngine 'render.ps1' $args2 (Join-Path $job 'render')
  $ok = ($log -match '(?m)^OK ')
  $pngs = @(); if (Test-Path (Join-Path $job 'png')) { $pngs = @(Get-ChildItem (Join-Path $job 'png') -Filter 'slide-*.png' | Sort-Object Name | ForEach-Object { "/png/$jobName/" + $_.Name }) }
  if ($ok -and $req.styleJson) {
    # layout history keeps the next recommendation different (shared with the varied-ppt skill)
    try {
      $st = ([string]$req.styleJson) | ConvertFrom-Json
      $hist = @(); if (Test-Path $HistoryPath) { $hist = @(FlatHistory ((ReadText $HistoryPath) | ConvertFrom-Json)) }
      $entry = [ordered]@{ date = (Get-Date -Format 'yyyy-MM-dd HH:mm'); deck = $name; style = [string]$st.name }
      foreach ($k in 'cover','coverTitle','mirror','titleAlign','kicker','motif','contentBg','agenda','cards','stats','process') { $entry[$k] = $st.$k }
      if ($req.brand) { $entry.dk2 = [string]$req.brand.dk2; $entry.accent1 = [string]$req.brand.accent1; $entry.headFont = [string]$req.brand.headFont }
      $hist += [pscustomobject]$entry
      if ($hist.Count -gt 30) { $hist = $hist[($hist.Count - 30)..($hist.Count - 1)] }
      WriteText $HistoryPath (ConvertTo-Json -InputObject @($hist) -Depth 4)
    } catch { Write-Host "[render] history: $_" }
  }
  Write-Host ("[render] " + $(if ($ok) { 'OK ' + $out } else { 'FAILED' }))
  SendJson $ctx ([ordered]@{ ok = [bool]$ok; out = $out; potx = $(if ($req.potx) { [IO.Path]::ChangeExtension($out, '.potx') } else { $null }); pdf = $(if ($req.pdf) { [IO.Path]::ChangeExtension($out, '.pdf') } else { $null }); pngs = $pngs; log = $log })
}
function ImportDeck($ctx) {
  $orig = QS $ctx 'name'
  $ext = [IO.Path]::GetExtension($orig).ToLower()
  if (@('.pptx', '.ppt', '.potx', '.pptm') -notcontains $ext) { SendJson $ctx @{ error = "PowerPoint 파일(.pptx, .ppt)만 가져올 수 있어요." } 400; return }
  $bytes = BodyBytes $ctx
  if ($bytes.Length -gt 300MB) { SendJson $ctx @{ error = 'file too large (300MB max)' } 400; return }
  $id = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,6)
  $dir = Join-Path $ImpDir $id
  New-Item -ItemType Directory -Force $dir | Out-Null
  $srcFile = Join-Path $dir ('source' + $ext)
  [IO.File]::WriteAllBytes($srcFile, $bytes)
  WriteText (Join-Path $dir 'name.txt') ([IO.Path]::GetFileNameWithoutExtension($orig))
  Write-Host ("[import] " + $orig + " ...")
  $log = RunEngine 'import.ps1' @('-Source', $srcFile, '-OutDir', $dir) (Join-Path $dir 'import')
  if ($log -notmatch '(?m)^OK ' -or -not (Test-Path (Join-Path $dir 'info.json'))) { Write-Host '[import] FAILED'; SendJson $ctx @{ error = 'PowerPoint로 파일을 열지 못했어요.'; log = $log } 500; return }
  Write-Host '[import] OK'
  SendRaw $ctx ('{"id":' + (JStr $id) + ',"name":' + (JStr ([IO.Path]::GetFileNameWithoutExtension($orig))) + ',"dir":' + (JStr $dir) + ',"info":' + (ReadText (Join-Path $dir 'info.json')) + '}')
}
function ApplyEdits($ctx) {
  $req = (BodyText $ctx) | ConvertFrom-Json
  $id = [string]$req.id
  if ($id -notmatch '^[0-9a-f\-]+$') { SendJson $ctx @{ error = 'bad id' } 400; return }
  $dir = Join-Path $ImpDir $id
  $srcFile = Get-ChildItem $dir -Filter 'source.*' | Select-Object -First 1
  if (-not $srcFile) { SendJson $ctx @{ error = '원본 파일을 찾을 수 없어요. 다시 가져와 주세요.' } 404; return }
  $stamp = 'result-' + (Get-Date -Format 'HHmmss')
  $work = Join-Path $dir $stamp
  New-Item -ItemType Directory -Force $work | Out-Null
  WriteText (Join-Path $work 'edits.json') ([string]$req.editsJson)
  $out = UniqueOut (SafeName ([string]$req.fileName))
  $args2 = @('-Source', $srcFile.FullName, '-Edits', (Join-Path $work 'edits.json'), '-Out', $out, '-PngDir', (Join-Path $work 'png'))
  if ($req.brandJson) { WriteText (Join-Path $work 'brand.json') ([string]$req.brandJson); $args2 += @('-Brand', (Join-Path $work 'brand.json')) }
  Write-Host ("[apply] " + $out + " ...")
  $log = RunEngine 'apply.ps1' $args2 (Join-Path $work 'apply')
  $ok = ($log -match '(?m)^OK ')
  $pngs = @(); if (Test-Path (Join-Path $work 'png')) { $pngs = @(Get-ChildItem (Join-Path $work 'png') -Filter 'slide-*.png' | Sort-Object Name | ForEach-Object { "/imports/$id/$stamp/png/" + $_.Name }) }
  Write-Host ("[apply] " + $(if ($ok) { 'OK' } else { 'FAILED' }))
  SendJson $ctx ([ordered]@{ ok = [bool]$ok; out = $out; pngs = $pngs; log = $log })
}

# ---------- Start listener ----------
$listener = $null; $port = 0
foreach ($try in 8765..8785) {
  $l = New-Object Net.HttpListener
  $l.Prefixes.Add("http://localhost:$try/")
  try { $l.Start(); $listener = $l; $port = $try; break } catch { $l.Close() }
}
if (-not $listener) { Write-Host 'No free port between 8765 and 8785.'; exit 1 }
$url = "http://localhost:$port/"
Write-Host "PPT Editor running at $url  (close this window to stop)"
if (-not $env:PPT_EDITOR_NO_BROWSER) { Start-Process $url }

while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  try {
    $path = $ctx.Request.Url.AbsolutePath
    $method = $ctx.Request.HttpMethod
    $q = $ctx.Request.QueryString
    if ($path -eq '/' -or $path -eq '/index.html') { SendFile $ctx (Join-Path $AppDir 'app\index.html') }
    elseif ($path -eq '/api/state') {
      $projects = ConvertTo-Json -InputObject @(Get-ChildItem $ProjDir -Filter '*.json' | Sort-Object LastWriteTime -Descending | ForEach-Object { $_.BaseName }) -Compress
      $json = '{"fonts":' + $FontsJson +
        ',"layouts":' + (RawOr (Join-Path $EngineDir 'layouts.json') '[]') +
        ',"history":' + (RawOr $HistoryPath '[]') +
        ',"brands":' + (RawOr $BrandsPath 'null') +
        ',"projects":' + $projects +
        ',"outDir":' + (JStr $OutDir) +
        ',"engineDir":' + (JStr $EngineDir) + '}'
      SendRaw $ctx $json
    }
    elseif ($path -eq '/api/brands' -and $method -eq 'POST') { WriteText $BrandsPath (BodyText $ctx); SendJson $ctx @{ ok = $true } }
    elseif ($path -eq '/api/project' -and $method -eq 'GET') {
      $f = Join-Path $ProjDir ((SafeName (QS $ctx 'name')) + '.json'); if (Test-Path -LiteralPath $f) { SendRaw $ctx (ReadText $f) } else { SendJson $ctx @{ error = 'not found' } 404 }
    }
    elseif ($path -eq '/api/project' -and $method -eq 'POST') { WriteText (Join-Path $ProjDir ((SafeName (QS $ctx 'name')) + '.json')) (BodyText $ctx); SendJson $ctx @{ ok = $true } }
    elseif ($path -eq '/api/fontalias') { SendRaw $ctx (FontAlias) }
    elseif ($path -eq '/api/upload' -and $method -eq 'POST') { Upload $ctx }
    elseif ($path -eq '/api/build' -and $method -eq 'POST') { Build $ctx }
    elseif ($path -eq '/api/render' -and $method -eq 'POST') { RenderDoc $ctx }
    elseif ($path -match '^/js/([a-z]+\.js)$') { SendFile $ctx (Join-Path $AppDir ('app\js\' + $Matches[1])) }
    elseif ($path -eq '/api/import' -and $method -eq 'POST') { ImportDeck $ctx }
    elseif ($path -eq '/api/apply' -and $method -eq 'POST') { ApplyEdits $ctx }
    elseif ($path -match '^/imports/([0-9a-f\-]+)/((result-\d+/)?(png|media))/([A-Za-z0-9_\-]+\.png)$') { SendFile $ctx (Join-Path $ImpDir ($Matches[1] + '\' + $Matches[2].Replace('/', '\') + '\' + $Matches[5])) }
    elseif ($path -eq '/api/open' -and $method -eq 'POST') {
      $req = (BodyText $ctx) | ConvertFrom-Json
      $target = [string]$req.path
      if ($req.what -eq 'folder') {
        if ($target -and (IsUnder $target $OutDir) -and (Test-Path -LiteralPath $target)) { Start-Process explorer.exe ('/select,"' + $target + '"') } else { Start-Process explorer.exe ('"' + $OutDir + '"') }
        SendJson $ctx @{ ok = $true }
      } elseif ($target -and (IsUnder $target $OutDir)) { Start-Process -FilePath $target; SendJson $ctx @{ ok = $true } }
      else { SendJson $ctx @{ error = 'outside output folder' } 403 }
    }
    elseif ($path -match '^/png/([0-9\-]+)/(slide-\d+\.png)$') { SendFile $ctx (Join-Path $WorkRoot ($Matches[1] + '\png\' + $Matches[2])) }
    elseif ($path -match '^/assets/([A-Za-z0-9_\-]+\.(svg|png))$') { SendFile $ctx (Join-Path $EngineDir ('assets\' + $Matches[1])) }
    elseif ($path -match '^/uploads/([A-Za-z0-9_\-]+\.(png|jpe?g|gif|bmp|svg))$') { SendFile $ctx (Join-Path $UpDir $Matches[1]) }
    else { Send $ctx 404 'text/plain' ($utf8.GetBytes('not found')) }
  } catch {
    Write-Host ("[error] " + $_)
    try { SendJson $ctx @{ error = "$_" } 500 } catch {}
  }
}
