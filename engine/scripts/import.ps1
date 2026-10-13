param(
  [Parameter(Mandatory=$true)][string]$Source,
  [Parameter(Mandatory=$true)][string]$OutDir
)
# Reads an existing .pptx/.ppt with PowerPoint (COM) and writes OutDir\info.json:
# slide size, per-slide PNG previews, and every shape (text paragraphs, pictures, tables, charts) addressed by a key
# ("3" = 3rd shape, "3.2" = 2nd item of group 3) that apply.ps1 uses to write edits back.
$ErrorActionPreference = 'Stop'
function LongPath([string]$p) { $full = [IO.Path]::GetFullPath($p); if ($full.Length -gt 240 -and -not $full.StartsWith('\\?\')) { return '\\?\' + $full }; return $full }

$tmp = Join-Path $env:TEMP ('ppt-import-' + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Force $tmp | Out-Null
$src = Join-Path $tmp ('source' + [IO.Path]::GetExtension($Source))
Copy-Item -LiteralPath (LongPath $Source) -Destination $src -Force
New-Item -ItemType Directory -Force (LongPath (Join-Path $OutDir 'png')), (LongPath (Join-Path $OutDir 'media')) | Out-Null

function Round1([double]$v) { return [math]::Round($v, 1) }
function ParaList($tr) {
  $list = @()
  $n = $tr.Paragraphs().Count
  for ($i = 1; $i -le $n; $i++) {
    $p = $tr.Paragraphs($i)
    $t = ([string]$p.Text).TrimEnd("`r").Replace([string][char]11, [string][char]0x21B5)
    $size = 0; try { $size = [double]$p.Font.Size } catch {}
    $bold = $false; try { $bold = ($p.Font.Bold -eq -1) } catch {}
    $list += [ordered]@{ text = $t; level = [int]$p.IndentLevel; size = $size; bold = $bold }
  }
  return ,$list
}

$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($src, -1, 0, 0)
  $W = [double]$pres.PageSetup.SlideWidth; $H = [double]$pres.PageSetup.SlideHeight
  $script:media = 0
  $slides = @()

  function Walk($coll, [string]$prefix, $list, [bool]$inGroup) {
    $count = $coll.Count
    for ($k = 1; $k -le $count; $k++) {
      $sh = $coll.Item($k)
      $key = if ($prefix) { "$prefix.$k" } else { "$k" }
      $o = [ordered]@{ key = $key; name = [string]$sh.Name; x = (Round1 $sh.Left); y = (Round1 $sh.Top); w = (Round1 $sh.Width); h = (Round1 $sh.Height); kind = 'other'; inGroup = $inGroup }
      $type = [int]$sh.Type
      if ($type -eq 14) { try { $o.ph = [int]$sh.PlaceholderFormat.Type } catch {} }
      if ($type -eq 6) { $o.kind = 'group'; $list.Add($o) | Out-Null; Walk $sh.GroupItems $key $list $true; continue }
      $isPicture = ($type -eq 13 -or $type -eq 11)
      if ($type -eq 14) { try { if ([int]$sh.PlaceholderFormat.ContainedType -eq 13) { $isPicture = $true } } catch {} }
      if ($sh.HasTable -eq -1) {
        $o.kind = 'table'
        $rows = @()
        for ($r = 1; $r -le $sh.Table.Rows.Count; $r++) {
          $row = @(); for ($c = 1; $c -le $sh.Table.Columns.Count; $c++) { $row += ([string]$sh.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text).Replace("`r", ' ').Replace([string][char]11, ' ') }
          $rows += ,$row
        }
        $o.table = $rows
      } elseif ($sh.HasChart -eq -1) {
        $o.kind = 'chart'
        try {
          $ch = $sh.Chart; $o.chartType = [int]$ch.ChartType
          $s1 = $ch.SeriesCollection(1)
          $o.seriesName = [string]$s1.Name
          $o.categories = @($s1.XValues | ForEach-Object { [string]$_ })
          $o.values = @($s1.Values | ForEach-Object { [double]$_ })
        } catch {}
      } elseif ($isPicture) {
        $o.kind = 'picture'
        $script:media++
        $name = 'img-{0:D3}.png' -f $script:media
        try {
          $sh.Export((Join-Path $tmp $name), 2)
          Copy-Item -LiteralPath (Join-Path $tmp $name) -Destination (LongPath (Join-Path $OutDir "media\$name")) -Force
          $o.image = $name
        } catch {}
      }
      if ($o.kind -eq 'other' -or $o.kind -eq 'picture') {
        $hasText = $false
        try { $hasText = ($sh.HasTextFrame -eq -1) -and ($sh.TextFrame.HasText -eq -1) } catch {}
        if ($hasText -and $o.kind -eq 'other') { $o.kind = 'text'; $o.paragraphs = (ParaList $sh.TextFrame.TextRange) }
        elseif ($o.kind -eq 'other') { try { if ($sh.HasTextFrame -eq -1 -and $type -eq 14) { $o.kind = 'text'; $o.paragraphs = @([ordered]@{ text = ''; level = 1; size = 0; bold = $false }); $o.empty = $true } } catch {} }
      }
      $list.Add($o) | Out-Null
    }
  }

  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $png = 'slide-{0:D2}.png' -f $i
    $sl.Export((Join-Path $tmp $png), 'PNG', 1280, [int](1280 * $H / $W))
    Copy-Item -LiteralPath (Join-Path $tmp $png) -Destination (LongPath (Join-Path $OutDir "png\$png")) -Force
    $list = New-Object System.Collections.ArrayList
    Walk $sl.Shapes '' $list $false
    $notes = ''
    try { $notes = ([string]$sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text).Replace("`r", "`n") } catch {}
    $layout = ''; try { $layout = [string]$sl.CustomLayout.Name } catch {}
    $slides += [ordered]@{ index = $i; png = $png; layout = $layout; hidden = ($sl.SlideShowTransition.Hidden -eq -1); notes = $notes; shapes = @($list) }
  }
  $info = [ordered]@{ width = $W; height = $H; slideCount = $pres.Slides.Count; slides = $slides }
  $pres.Close()
  [IO.File]::WriteAllText((LongPath (Join-Path $OutDir 'info.json')), (ConvertTo-Json -InputObject $info -Depth 12 -Compress), (New-Object Text.UTF8Encoding $false))
  Write-Output "OK $($info.slideCount)"
}
catch { Write-Output ("IMPORT ERROR: " + $_ + " @ " + $_.ScriptStackTrace); exit 1 }
finally {
  # Never close the user's own PowerPoint: quit only when no other presentation is open.
  $quit = $false
  try { if ($app.Presentations.Count -eq 0) { $app.Quit(); $quit = $true } } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  # Wait for PowerPoint to exit so the next job does not attach to an instance that is shutting down.
  if ($quit) { for ($w = 0; $w -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Milliseconds 500 } }
}
