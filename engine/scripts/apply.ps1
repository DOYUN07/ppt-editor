param(
  [Parameter(Mandatory=$true)][string]$Source,
  [Parameter(Mandatory=$true)][string]$Edits,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$Brand,
  [string]$PngDir
)
# Applies edits to an existing presentation while keeping its own design, and saves a copy.
# Edits JSON: { "slides": [ { "index": 3, "delete": false, "notes": "...",
#   "shapes": [ { "key": "2", "paragraphs": ["...", "..."] }, { "key": "5", "image": "C:\\path.png" }, { "key": "4", "table": [["a","b"]] } ] } ] }
# Shape keys come from import.ps1 ("3" = 3rd shape on the slide, "3.2" = 2nd item inside group 3).
# With -Brand, the theme colors and fonts of every slide master are replaced by the brand's.
$ErrorActionPreference = 'Stop'
function LongPath([string]$p) { $full = [IO.Path]::GetFullPath($p); if ($full.Length -gt 240 -and -not $full.StartsWith('\\?\')) { return '\\?\' + $full }; return $full }
function ReadJson([string]$p) { return ([IO.File]::ReadAllText((LongPath $p), [Text.Encoding]::UTF8) | ConvertFrom-Json) }
function Has($o, [string]$name) { return ($null -ne $o) -and ($o.PSObject.Properties.Name -contains $name) -and ($null -ne $o.$name) }
function RGBc([string]$hex) { $hex = $hex.TrimStart('#'); return [int]([Convert]::ToInt32($hex.Substring(0,2),16) + [Convert]::ToInt32($hex.Substring(2,2),16)*256 + [Convert]::ToInt32($hex.Substring(4,2),16)*65536) }

$ed = ReadJson $Edits
$tmp = Join-Path $env:TEMP ('ppt-apply-' + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Force $tmp | Out-Null
$src = Join-Path $tmp ('source' + [IO.Path]::GetExtension($Source))
Copy-Item -LiteralPath (LongPath $Source) -Destination $src -Force
$SOFT = [string][char]0x21B5

function ByKey($slide, [string]$key) {
  $parts = $key.Split('.')
  $sh = $slide.Shapes.Item([int]$parts[0])
  for ($i = 1; $i -lt $parts.Count; $i++) { $sh = $sh.GroupItems.Item([int]$parts[$i]) }
  return $sh
}
function SetParagraphs($sh, $paras) {
  $tr = $sh.TextFrame.TextRange
  $new = @($paras | ForEach-Object { ([string]$_).Replace($SOFT, [string][char]11) })
  $count = $tr.Paragraphs().Count
  if ($new.Count -eq $count) {
    # Same paragraph count: replace each paragraph's characters so its formatting and indent level stay.
    for ($i = 1; $i -le $count; $i++) {
      $p = $tr.Paragraphs($i)
      $old = ([string]$p.Text).TrimEnd("`r")
      if ($old -ceq $new[$i-1]) { continue }
      if ($old.Length -gt 0) { $p.Characters(1, $old.Length).Text = $new[$i-1] }
      else { $null = $p.Characters(1, 0).InsertAfter($new[$i-1]) }
    }
  } else {
    $tr.Text = ($new -join "`r")
  }
}
function Stage([string]$p) { $dst = Join-Path $tmp ('img-' + [guid]::NewGuid().ToString('N').Substring(0,6) + [IO.Path]::GetExtension($p)); Copy-Item -LiteralPath (LongPath $p) -Destination $dst -Force; return $dst }
function ReplacePicture($slide, $old, [string]$path) {
  $l = [double]$old.Left; $t = [double]$old.Top; $w = [double]$old.Width; $h = [double]$old.Height
  $pic = $slide.Shapes.AddPicture((Stage $path), 0, -1, [single]$l, [single]$t)
  $pic.LockAspectRatio = 0
  $w0 = [double]$pic.Width; $h0 = [double]$pic.Height
  $r = [math]::Max($w / $w0, $h / $h0)
  $cx = ($w0 - $w / $r) / 2; $cy = ($h0 - $h / $r) / 2
  $pic.PictureFormat.CropLeft = [single]$cx; $pic.PictureFormat.CropRight = [single]$cx
  $pic.PictureFormat.CropTop = [single]$cy; $pic.PictureFormat.CropBottom = [single]$cy
  $pic.Left = [single]$l; $pic.Top = [single]$t; $pic.Width = [single]$w; $pic.Height = [single]$h
  $target = [int]$old.ZOrderPosition
  $guard = 0
  while ([int]$pic.ZOrderPosition -gt $target + 1 -and $guard -lt 500) { $pic.ZOrder(3); $guard++ }   # msoSendBackward
  $pic.Name = $old.Name
  $old.Delete()
}

$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($src, 0, 0, 0)
  $deletes = @()
  foreach ($se in @($ed.slides)) {
    $idx = [int]$se.index
    if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) { continue }
    $slide = $pres.Slides.Item($idx)
    if ((Has $se 'delete') -and $se.delete) { $deletes += $idx; continue }
    if (Has $se 'notes') { try { $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = ([string]$se.notes).Replace("`n", "`r") } catch { Write-Output "WARN notes slide $idx : $_" } }
    if (-not (Has $se 'shapes')) { continue }
    # Pictures last and in reverse key order so deleting/adding shapes cannot shift the keys of later edits.
    $shapeEdits = @($se.shapes)
    foreach ($e in ($shapeEdits | Where-Object { -not (Has $_ 'image') })) {
      try {
        $sh = ByKey $slide ([string]$e.key)
        if (Has $e 'paragraphs') { SetParagraphs $sh @($e.paragraphs) }
        if (Has $e 'table') {
          $rows = @($e.table)
          for ($r = 1; $r -le [math]::Min($rows.Count, $sh.Table.Rows.Count); $r++) {
            $row = @($rows[$r-1])
            for ($c = 1; $c -le [math]::Min($row.Count, $sh.Table.Columns.Count); $c++) {
              $tr = $sh.Table.Cell($r, $c).Shape.TextFrame.TextRange
              if (([string]$tr.Text).Replace("`r", ' ').Replace([string][char]11, ' ') -cne [string]$row[$c-1]) { $tr.Text = [string]$row[$c-1] }
            }
          }
        }
      } catch { Write-Output "WARN slide $idx shape $($e.key): $_" }
    }
    $pics = @($shapeEdits | Where-Object { Has $_ 'image' } | Sort-Object { [int](([string]$_.key).Split('.')[0]) } -Descending)
    foreach ($e in $pics) {
      try { ReplacePicture $slide (ByKey $slide ([string]$e.key)) ([string]$e.image) } catch { Write-Output "WARN slide $idx picture $($e.key): $_" }
    }
  }
  foreach ($d in ($deletes | Sort-Object -Descending)) { $pres.Slides.Item($d).Delete() }

  if ($Brand) {
    $br = ReadJson $Brand
    $order = @('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
    for ($m = 1; $m -le $pres.Designs.Count; $m++) {
      $theme = $pres.Designs.Item($m).SlideMaster.Theme
      for ($i = 0; $i -lt 12; $i++) { $theme.ThemeColorScheme.Colors($i+1).RGB = (RGBc ([string]$br.colors.($order[$i]))) }
      if (Has $br 'headFont') { $theme.ThemeFontScheme.MajorFont.Item(1).Name = [string]$br.headFont; $theme.ThemeFontScheme.MajorFont.Item(3).Name = [string]$br.headFont }
      if (Has $br 'bodyFont') { $theme.ThemeFontScheme.MinorFont.Item(1).Name = [string]$br.bodyFont; $theme.ThemeFontScheme.MinorFont.Item(3).Name = [string]$br.bodyFont }
    }
  }

  $name = [IO.Path]::GetFileNameWithoutExtension($Out)
  $tout = Join-Path $tmp ($name + '.pptx')
  $pres.SaveAs($tout, 24)
  if ($PngDir) {
    New-Item -ItemType Directory -Force (LongPath $PngDir) | Out-Null
    $W = [double]$pres.PageSetup.SlideWidth; $H = [double]$pres.PageSetup.SlideHeight
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $png = Join-Path $tmp ("slide-{0:D2}.png" -f $i)
      $pres.Slides.Item($i).Export($png, 'PNG', 1280, [int](1280 * $H / $W))
      Copy-Item -LiteralPath $png -Destination (LongPath (Join-Path $PngDir ("slide-{0:D2}.png" -f $i))) -Force
    }
  }
  $pres.Close()
  $outFull = LongPath $Out
  New-Item -ItemType Directory -Force (Split-Path $outFull -Parent) | Out-Null
  Copy-Item -LiteralPath $tout -Destination $outFull -Force
  Write-Output "OK $Out"
}
catch { Write-Output ("APPLY ERROR: " + $_ + " @ " + $_.ScriptStackTrace); try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}; exit 1 }
finally {
  # Never close the user's own PowerPoint: quit only when no other presentation is open.
  $quit = $false
  try { if ($app.Presentations.Count -eq 0) { $app.Quit(); $quit = $true } } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  # Wait for PowerPoint to exit so the next job does not attach to an instance that is shutting down.
  if ($quit) { for ($w = 0; $w -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Milliseconds 500 } }
}
