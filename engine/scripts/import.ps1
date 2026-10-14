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
Add-Type -AssemblyName System.Drawing
# pixel size of an exported shape image; rotated shapes export their rotated outline, so the editor needs the real size
function PxSize([string]$file) { try { $im = [Drawing.Image]::FromFile($file); $r = @($im.Width, $im.Height); $im.Dispose(); return ,$r } catch { return $null } }
function Hex([int]$v) { return ('{0:X2}{1:X2}{2:X2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255)) }
# Shape types the editor can rebuild natively (MsoAutoShapeType); anything else is kept as a picture snapshot.
$KNOWN_AST = @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 18, 21, 22, 33, 34, 35, 36, 37, 51, 52, 91, 92, 106, 147, 152, 155, 158, 179)
function StyleOf($sh, [int]$type) {
  $st = [ordered]@{}
  try { $r = [double]$sh.Rotation; if ($r -ne 0) { $st.rot = $r } } catch {}
  try { if ($sh.HorizontalFlip -eq -1) { $st.flipH = $true }; if ($sh.VerticalFlip -eq -1) { $st.flipV = $true } } catch {}
  if ($type -eq 1 -or $type -eq 17 -or $type -eq 14) { try { $st.ast = [int]$sh.AutoShapeType; if ($st.ast -eq 5 -or $st.ast -eq 152 -or $st.ast -eq 155) { $st.radius = [double]$sh.Adjustments.Item(1) } } catch {} }
  try { if ($sh.Fill.Visible -eq -1) { if ($sh.Fill.Type -eq 1) { $st.fill = Hex $sh.Fill.ForeColor.RGB; $tr = [double]$sh.Fill.Transparency; if ($tr -gt 0) { $st.fillAlpha = [math]::Round(1 - $tr, 2) } } else { $st.complexFill = $true } } } catch {}
  try { if ($sh.Line.Visible -eq -1) { $st.line = [ordered]@{ color = (Hex $sh.Line.ForeColor.RGB); width = [double]$sh.Line.Weight; dash = [int]$sh.Line.DashStyle; endArrow = [int]$sh.Line.EndArrowheadStyle; beginArrow = [int]$sh.Line.BeginArrowheadStyle } } } catch {}
  try { if ($sh.HasTextFrame -eq -1) { $tf = $sh.TextFrame; $st.valign = [int]$tf.VerticalAnchor; $st.pad = @([double]$tf.MarginTop, [double]$tf.MarginRight, [double]$tf.MarginBottom, [double]$tf.MarginLeft); if ($tf.WordWrap -eq 0) { $st.nowrap = $true } } } catch {}
  try { if ($sh.Shadow.Visible -eq -1) { $st.shadow = $true } } catch {}
  return $st
}
function ParaList($tr) {
  $list = @()
  $n = $tr.Paragraphs().Count
  for ($i = 1; $i -le $n; $i++) {
    $p = $tr.Paragraphs($i)
    $t = ([string]$p.Text).TrimEnd("`r").Replace([string][char]11, [string][char]0x21B5)
    $size = 0; try { $size = [double]$p.Font.Size } catch {}
    $bold = $false; try { $bold = ($p.Font.Bold -eq -1) } catch {}
    $runs = @()
    try {
      $rn = $p.Runs().Count
      for ($j = 1; $j -le $rn; $j++) {
        $r = $p.Runs($j); $rt = ([string]$r.Text).TrimEnd("`r").Replace([string][char]11, "`n")
        if ($rt.Length -eq 0) { continue }
        $ro = [ordered]@{ t = $rt }
        try { $ro.size = [double]$r.Font.Size } catch {}
        try { if ($r.Font.Bold -eq -1) { $ro.b = $true } } catch {}
        try { if ($r.Font.Italic -eq -1) { $ro.i = $true } } catch {}
        try { if ($r.Font.Underline -eq -1) { $ro.u = $true } } catch {}
        try { $ro.color = Hex $r.Font.Color.RGB } catch {}
        try { $ro.font = [string]$r.Font.Name } catch {}
        $runs += $ro
      }
    } catch {}
    $align = 1; try { $align = [int]$p.ParagraphFormat.Alignment } catch {}
    $bullet = 0; try { if ($p.ParagraphFormat.Bullet.Visible -eq -1) { $bullet = [int]$p.ParagraphFormat.Bullet.Type } } catch {}
    $po = [ordered]@{ text = $t; level = [int]$p.IndentLevel; size = $size; bold = $bold; runs = $runs; align = $align; bullet = $bullet }
    if ($bullet -eq 1) { try { $po.bulletChar = [int]$p.ParagraphFormat.Bullet.Character; $po.bulletFont = [string]$p.ParagraphFormat.Bullet.Font.Name } catch {} }
    # spacing: *Rule -1 = multiples of a line, otherwise points
    try { $pf = $p.ParagraphFormat
      $po.lsRule = [int]$pf.LineRuleWithin; $po.ls = [double]$pf.SpaceWithin
      $po.sbRule = [int]$pf.LineRuleBefore; $po.sb = [double]$pf.SpaceBefore
      $po.saRule = [int]$pf.LineRuleAfter; $po.sa = [double]$pf.SpaceAfter } catch {}
    $list += $po
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
        # look: bar colour, number format, axis / gridlines, data labels
        try { $o.seriesColor = Hex $s1.Format.Fill.ForeColor.RGB } catch {}
        try { $pc = @(); for ($pi = 1; $pi -le @($o.values).Count; $pi++) { $pc += (Hex $s1.Points($pi).Format.Fill.ForeColor.RGB) }; if (@($pc | Select-Object -Unique).Count -gt 1) { $o.pointColors = $pc } } catch {}
        try { if ($s1.HasDataLabels) { $o.labels = $true; $o.numFmt = [string]$s1.DataLabels().NumberFormat } else { $o.labels = $false } } catch {}
        try { if (-not $o.numFmt -or $o.numFmt -eq 'General') { $o.numFmt = [string]$ch.Axes(2).TickLabels.NumberFormat } } catch {}
        try { $va = $ch.Axes(2); if (-not $va.MinimumScaleIsAuto) { $o.axisMin = [double]$va.MinimumScale }; if (-not $va.MaximumScaleIsAuto) { $o.axisMax = [double]$va.MaximumScale } } catch {}
        try { $o.valueAxis = [bool]($ch.HasAxis(2) -and $ch.Axes(2).TickLabelPosition -ne -4142); $o.gridlines = [bool]$ch.Axes(2).HasMajorGridlines } catch {}
      } elseif ($isPicture) {
        $o.kind = 'picture'
        $script:media++
        $name = 'img-{0:D3}.png' -f $script:media
        try {
          $sh.Export((Join-Path $tmp $name), 2)
          Copy-Item -LiteralPath (Join-Path $tmp $name) -Destination (LongPath (Join-Path $OutDir "media\$name")) -Force
          $o.image = $name; $o.px = (PxSize (Join-Path $tmp $name))
        } catch {}
      }
      if ($o.kind -eq 'other' -or $o.kind -eq 'picture') {
        $hasText = $false
        try { $hasText = ($sh.HasTextFrame -eq -1) -and ($sh.TextFrame.HasText -eq -1) } catch {}
        if ($hasText -and $o.kind -eq 'other') { $o.kind = 'text'; $o.paragraphs = (ParaList $sh.TextFrame.TextRange) }
        elseif ($o.kind -eq 'other') { try { if ($sh.HasTextFrame -eq -1 -and $type -eq 14) { $o.kind = 'text'; $o.paragraphs = @([ordered]@{ text = ''; level = 1; size = 0; bold = $false }); $o.empty = $true } } catch {} }
      }
      $o.style = (StyleOf $sh $type)
      $needSnap = ($o.kind -eq 'other' -and $type -ne 9 -and -not ($type -eq 1 -and $KNOWN_AST -contains $o.style.ast)) -or $o.style.complexFill -or ($type -eq 1 -and -not ($KNOWN_AST -contains $o.style.ast))
      if ($type -eq 9) { $o.kind = 'line' }
      if ($needSnap -and $o.kind -ne 'picture' -and $o.kind -ne 'table' -and $o.kind -ne 'chart') {
        $script:media++
        $sn = 'snap-{0:D3}.png' -f $script:media
        try { $sh.Export((Join-Path $tmp $sn), 2); Copy-Item -LiteralPath (Join-Path $tmp $sn) -Destination (LongPath (Join-Path $OutDir "media\$sn")) -Force; $o.snap = $sn; $o.px = (PxSize (Join-Path $tmp $sn)) } catch {}
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
    $bg = 'bg-{0:D2}.png' -f $i
    try {
      $dup = $sl.Duplicate().Item(1)
      while ($dup.Shapes.Count -gt 0) { $dup.Shapes.Item(1).Delete() }
      $dup.Export((Join-Path $tmp $bg), 'PNG', 1600, [int](1600 * $H / $W))
      $dup.Delete()
      Copy-Item -LiteralPath (Join-Path $tmp $bg) -Destination (LongPath (Join-Path $OutDir "media\$bg")) -Force
    } catch { $bg = $null }
    $slides += [ordered]@{ index = $i; png = $png; bg = $bg; layout = $layout; hidden = ($sl.SlideShowTransition.Hidden -eq -1); notes = $notes; shapes = @($list) }
  }
  $info = [ordered]@{ width = $W; height = $H; slideCount = $pres.Slides.Count; slides = $slides }
  $pres.Close()
  [IO.File]::WriteAllText((LongPath (Join-Path $OutDir 'info.json')), (ConvertTo-Json -InputObject $info -Depth 12 -Compress), (New-Object Text.UTF8Encoding $false))
  Write-Output "OK $($info.slideCount)"
}
catch { Write-Output ("IMPORT ERROR: " + $_ + " @ " + $_.ScriptStackTrace); try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}; exit 1 }
finally {
  # Never close the user's own PowerPoint: quit only when no other presentation is open.
  $quit = $false
  try { if ($app.Presentations.Count -eq 0) { $app.Quit(); $quit = $true } } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  # Wait for PowerPoint to exit so the next job does not attach to an instance that is shutting down.
  if ($quit) { for ($w = 0; $w -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Milliseconds 500 } }
}
