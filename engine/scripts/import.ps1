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
# Difference matting: the same shape rendered on black and on white gives exact colour + transparency
# (soft shadows, gradients, partial arcs), cropped tight to what is actually drawn.
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System; using System.Drawing; using System.Drawing.Imaging; using System.Runtime.InteropServices;
public static class Matte {
  static byte[] Read(Bitmap b, out int stride) {
    var d = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
    stride = d.Stride; var a = new byte[stride * b.Height]; Marshal.Copy(d.Scan0, a, 0, a.Length); b.UnlockBits(d); return a;
  }
  // returns "x,y,w,h" (pixels in the slide image) of the saved cut-out, or "" when nothing is visible
  public static string Cut(string blackPng, string whitePng, string outPng, int x0, int y0, int x1, int y1) {
    using (var bk = new Bitmap(blackPng)) using (var wt = new Bitmap(whitePng)) {
      int W = bk.Width, H = bk.Height, sb, sw;
      x0 = Math.Max(0, x0); y0 = Math.Max(0, y0); x1 = Math.Min(W, x1); y1 = Math.Min(H, y1);
      if (x1 <= x0 || y1 <= y0) return "";
      byte[] B = Read(bk, out sb), Wt = Read(wt, out sw);
      int rw = x1 - x0, rh = y1 - y0; var A = new byte[rw * rh]; var C = new byte[rw * rh * 4];
      int minX = rw, minY = rh, maxX = -1, maxY = -1;
      for (int y = 0; y < rh; y++) for (int x = 0; x < rw; x++) {
        int p = (y + y0) * sb + (x + x0) * 4, q = (y + y0) * sw + (x + x0) * 4;
        int diff = ((Wt[q] - B[p]) + (Wt[q + 1] - B[p + 1]) + (Wt[q + 2] - B[p + 2])) / 3;
        int al = Math.Max(0, Math.Min(255, 255 - diff)); if (al < 4) al = 0;
        int o = (y * rw + x) * 4;
        if (al > 0) {
          for (int c = 0; c < 3; c++) C[o + c] = (byte)Math.Min(255, B[p + c] * 255 / al);
          if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y;
        }
        C[o + 3] = (byte)al;
      }
      if (maxX < 0) return "";
      int ow = maxX - minX + 1, oh = maxY - minY + 1;
      using (var outB = new Bitmap(ow, oh, PixelFormat.Format32bppArgb)) {
        var d = outB.LockBits(new Rectangle(0, 0, ow, oh), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
        var row = new byte[ow * 4];
        for (int y = 0; y < oh; y++) { Buffer.BlockCopy(C, ((y + minY) * rw + minX) * 4, row, 0, ow * 4); Marshal.Copy(row, 0, d.Scan0 + y * d.Stride, ow * 4); }
        outB.UnlockBits(d); outB.Save(outPng, ImageFormat.Png);
      }
      return (x0 + minX) + "," + (y0 + minY) + "," + ow + "," + oh;
    }
  }
}
'@
$SNAP_PX = 1920   # slide render width used for cut-outs
function ShapeAt($slide, [string]$key) { $parts = $key.Split('.'); $sh = $slide.Shapes.Item([int]$parts[0]); for ($i = 1; $i -lt $parts.Count; $i++) { $sh = $sh.GroupItems.Item([int]$parts[$i]) }; return $sh }
# Render each listed shape alone (black, then white background) on a copy of the slide and cut it out.
function SnapShapes($sl, $snaps, [double]$W, [double]$H) {
  if (-not $snaps.Count) { return }
  $dup = $sl.Duplicate().Item(1)
  try {
    $dup.FollowMasterBackground = 0; $dup.DisplayMasterShapes = 0
    $pxH = [int]($SNAP_PX * $H / $W); $k = $SNAP_PX / $W
    $top = @(); for ($i = 1; $i -le $dup.Shapes.Count; $i++) { $top += $dup.Shapes.Item($i) }
    foreach ($s in $snaps) {
      $o = $s.o; $rendered = $false
      try {
        foreach ($t in $top) { $t.Visible = 0 }
        # show the target and its parent groups, hide its siblings inside those groups
        $parts = $o.key.Split('.'); $cur = $dup.Shapes.Item([int]$parts[0]); $cur.Visible = -1
        for ($i = 1; $i -lt $parts.Count; $i++) {
          for ($g = 1; $g -le $cur.GroupItems.Count; $g++) { $cur.GroupItems.Item($g).Visible = 0 }
          $cur = $cur.GroupItems.Item([int]$parts[$i]); $cur.Visible = -1
        }
        $bk = Join-Path $tmp 'snap-black.png'; $wt = Join-Path $tmp 'snap-white.png'
        $dup.Background.Fill.Solid(); $dup.Background.Fill.ForeColor.RGB = 0; $dup.Export($bk, 'PNG', $SNAP_PX, $pxH)
        $dup.Background.Fill.ForeColor.RGB = 16777215; $dup.Export($wt, 'PNG', $SNAP_PX, $pxH)
        # search area: the rotated frame plus room for shadows / glow
        $rad = [math]::PI * [double]($(if ($o.style.rot) { $o.style.rot } else { 0 })) / 180
        $bw = [math]::Abs($o.w * [math]::Cos($rad)) + [math]::Abs($o.h * [math]::Sin($rad)); $bh = [math]::Abs($o.w * [math]::Sin($rad)) + [math]::Abs($o.h * [math]::Cos($rad))
        $cx = $o.x + $o.w / 2; $cy = $o.y + $o.h / 2; $padPt = 40
        $rendered = $true   # an empty result means the shape draws nothing (invisible frame): leave it out
        $res = [Matte]::Cut($bk, $wt, (Join-Path $tmp $s.file), [int](($cx - $bw / 2 - $padPt) * $k), [int](($cy - $bh / 2 - $padPt) * $k), [int](($cx + $bw / 2 + $padPt) * $k), [int](($cy + $bh / 2 + $padPt) * $k))
        if ($res) {
          $b = $res.Split(',') | ForEach-Object { [double]$_ }
          Copy-Item -LiteralPath (Join-Path $tmp $s.file) -Destination (LongPath (Join-Path $OutDir "media\$($s.file)")) -Force
          $o.snap = $s.file; $o.snapBox = @((Round1 ($b[0] / $k)), (Round1 ($b[1] / $k)), (Round1 ($b[2] / $k)), (Round1 ($b[3] / $k)))
        }
      } catch {}
      # fallback: PowerPoint's own shape export (cropped to the drawing, position approximate)
      if (-not $o.snap -and -not $rendered) { try { $sh = ShapeAt $sl $o.key; $sh.Export((Join-Path $tmp $s.file), 2); Copy-Item -LiteralPath (Join-Path $tmp $s.file) -Destination (LongPath (Join-Path $OutDir "media\$($s.file)")) -Force; $o.snap = $s.file; $o.px = (PxSize (Join-Path $tmp $s.file)) } catch {} }
    }
  } finally { try { $dup.Delete() } catch {} }
}
function Hex([int]$v) { return ('{0:X2}{1:X2}{2:X2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255)) }
# Shape types the editor can rebuild natively (MsoAutoShapeType); anything else is kept as a picture snapshot.
$KNOWN_AST = @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 18, 21, 22, 33, 34, 35, 36, 37, 51, 52, 91, 92, 106, 147, 152, 155, 158, 179)
function StyleOf($sh, [int]$type) {
  $st = [ordered]@{}
  try { $r = [double]$sh.Rotation; if ($r -ne 0) { $st.rot = $r } } catch {}
  try { if ($sh.HorizontalFlip -eq -1) { $st.flipH = $true }; if ($sh.VerticalFlip -eq -1) { $st.flipV = $true } } catch {}
  if ($type -eq 1 -or $type -eq 17 -or $type -eq 14) { try { $st.ast = [int]$sh.AutoShapeType; if ($st.ast -eq 5 -or $st.ast -eq 152 -or $st.ast -eq 155) { $st.radius = [double]$sh.Adjustments.Item(1) } } catch {} }
  try { if ($sh.Fill.Visible -eq -1) {
    $ft = [int]$sh.Fill.Type
    if ($ft -eq 1) { $st.fill = Hex $sh.Fill.ForeColor.RGB; $tr = [double]$sh.Fill.Transparency; if ($tr -gt 0) { $st.fillAlpha = [math]::Round(1 - $tr, 2) } }
    elseif ($ft -eq 3) {
      # gradient: stops + direction (linear angle, or radial from the centre)
      $f = $sh.Fill; $stops = @()
      for ($g = 1; $g -le $f.GradientStops.Count; $g++) { $gs = $f.GradientStops.Item($g); $stops += [ordered]@{ pos = [math]::Round([double]$gs.Position, 3); color = (Hex $gs.Color.RGB); alpha = [math]::Round(1 - [double]$gs.Transparency, 2) } }
      $gst = [int]$f.GradientStyle; $gr = [ordered]@{ type = $(if ($gst -eq 7 -or $gst -eq 5 -or $gst -eq 6) { 'radial' } else { 'linear' }); stops = $stops }
      try { $gr.angle = [double]$f.GradientAngle } catch { $gr.angle = 90 }
      $st.gradient = $gr; $st.fill = $stops[0].color
    }
    else { $st.complexFill = $true }   # pattern / texture / picture fill
  } } catch {}
  try { if ($sh.Line.Visible -eq -1) { $st.line = [ordered]@{ color = (Hex $sh.Line.ForeColor.RGB); width = [double]$sh.Line.Weight; dash = [int]$sh.Line.DashStyle; endArrow = [int]$sh.Line.EndArrowheadStyle; beginArrow = [int]$sh.Line.BeginArrowheadStyle } } } catch {}
  try { if ($sh.HasTextFrame -eq -1) { $tf = $sh.TextFrame; $st.valign = [int]$tf.VerticalAnchor; $st.pad = @([double]$tf.MarginTop, [double]$tf.MarginRight, [double]$tf.MarginBottom, [double]$tf.MarginLeft); if ($tf.WordWrap -eq 0) { $st.nowrap = $true } } } catch {}
  # effects kept as editable values: shadow (colour, blur, offset), glow, soft edge
  try { if ($sh.Shadow.Visible -eq -1) { $sd = $sh.Shadow; $st.shadow = $true
    $st.shadowFx = [ordered]@{ color = (Hex $sd.ForeColor.RGB); alpha = [math]::Round(1 - [double]$sd.Transparency, 2); blur = [math]::Round([double]$sd.Blur, 1); dx = [math]::Round([double]$sd.OffsetX, 1); dy = [math]::Round([double]$sd.OffsetY, 1) } } } catch {}
  try { if ([double]$sh.Glow.Radius -gt 0) { $st.glow = [ordered]@{ r = [math]::Round([double]$sh.Glow.Radius, 1); color = (Hex $sh.Glow.Color.RGB); alpha = [math]::Round(1 - [double]$sh.Glow.Transparency, 2) } } } catch {}
  try { if ([double]$sh.SoftEdge.Radius -gt 0) { $st.softEdge = [math]::Round([double]$sh.SoftEdge.Radius, 1) } } catch {}
  # only what the editor really cannot draw becomes a cut-out picture: reflection, 3-D,
  # and a shadow cast by a shape without fill (PowerPoint draws it from the invisible fill)
  try { if ([int]$sh.Reflection.Type -gt 0 -or $sh.ThreeD.Visible -eq -1) { $st.fx = $true } } catch {}
  if ($st.shadow -and -not $st.fill -and -not $st.gradient -and $type -ne 13 -and $type -ne 11) { $st.fx = $true }
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

# ---------- shape outlines as paths ----------
# DrawingML custom geometry (a:pathLst) -> "M x y L x y C ... Z" with coordinates normalised to the shape frame (0..1).
$INV = [Globalization.CultureInfo]::InvariantCulture
function Num([string]$v) { $d = 0.0; if (-not [double]::TryParse($v, [Globalization.NumberStyles]::Float, $INV, [ref]$d)) { throw "formula in path: $v" }; return $d }
function GeomFromPathLst($pathLst, $ns, [double]$extW, [double]$extH, $gbox, [double]$fw, [double]$fh) {
  $sb = New-Object Text.StringBuilder
  foreach ($path in $pathLst.SelectNodes('a:path', $ns)) {
    if ($path.GetAttribute('fill') -eq 'none') { continue }   # outline-only helper paths
    $pw = $(if ($path.GetAttribute('w')) { Num $path.GetAttribute('w') } else { $extW }); $ph = $(if ($path.GetAttribute('h')) { Num $path.GetAttribute('h') } else { $extH })
    if ($pw -le 0) { $pw = $extW }; if ($ph -le 0) { $ph = $extH }
    $norm = { param($x, $y)
      $nx = $x / $pw; $ny = $y / $ph
      if ($gbox) { $nx = ($gbox[0] + $nx * $gbox[2]) / $fw; $ny = ($gbox[1] + $ny * $gbox[3]) / $fh }
      return [string]::Format($INV, '{0:0.####} {1:0.####}', $nx, $ny) }
    $cx = 0.0; $cy = 0.0
    foreach ($c in $path.ChildNodes) {
      $pts = New-Object System.Collections.ArrayList
      foreach ($pt in $c.SelectNodes('a:pt', $ns)) { [void]$pts.Add([double[]]@((Num $pt.GetAttribute('x')), (Num $pt.GetAttribute('y')))) }
      switch ($c.LocalName) {
        'moveTo' { [void]$sb.Append('M' + (& $norm $pts[0][0] $pts[0][1]) + ' '); $cx = $pts[0][0]; $cy = $pts[0][1] }
        'lnTo' { [void]$sb.Append('L' + (& $norm $pts[0][0] $pts[0][1]) + ' '); $cx = $pts[0][0]; $cy = $pts[0][1] }
        'cubicBezTo' { [void]$sb.Append('C' + (($pts | ForEach-Object { & $norm $_[0] $_[1] }) -join ' ') + ' '); $cx = $pts[2][0]; $cy = $pts[2][1] }
        'quadBezTo' {
          $c1x = $cx + 2 / 3 * ($pts[0][0] - $cx); $c1y = $cy + 2 / 3 * ($pts[0][1] - $cy); $c2x = $pts[1][0] + 2 / 3 * ($pts[0][0] - $pts[1][0]); $c2y = $pts[1][1] + 2 / 3 * ($pts[0][1] - $pts[1][1])
          [void]$sb.Append('C' + (& $norm $c1x $c1y) + ' ' + (& $norm $c2x $c2y) + ' ' + (& $norm $pts[1][0] $pts[1][1]) + ' '); $cx = $pts[1][0]; $cy = $pts[1][1] }
        'arcTo' {
          $wR = Num $c.GetAttribute('wR'); $hR = Num $c.GetAttribute('hR'); $st = (Num $c.GetAttribute('stAng')) / 60000 * [math]::PI / 180; $sw = (Num $c.GetAttribute('swAng')) / 60000 * [math]::PI / 180
          # DrawingML angles are visual angles on the ellipse; convert to the ellipse parameter
          $par = { param($a) [math]::Atan2($wR * [math]::Sin($a), $hR * [math]::Cos($a)) }
          $t1 = & $par $st; $t2 = & $par ($st + $sw); $dt = $t2 - $t1
          if ([math]::Abs($sw) -ge 2 * [math]::PI - 1e-6) { $dt = [math]::Sign($sw) * 2 * [math]::PI } elseif ($sw -gt 0 -and $dt -le 0) { $dt += 2 * [math]::PI } elseif ($sw -lt 0 -and $dt -ge 0) { $dt -= 2 * [math]::PI }
          $ox = $cx - $wR * [math]::Cos($t1); $oy = $cy - $hR * [math]::Sin($t1)
          $n = [math]::Max(1, [math]::Ceiling([math]::Abs($dt) / ([math]::PI / 2))); $step = $dt / $n
          for ($s = 0; $s -lt $n; $s++) {
            $a = $t1 + $s * $step; $b = $a + $step; $k = 4 / 3 * [math]::Tan(($b - $a) / 4)
            $p1x = $ox + $wR * ([math]::Cos($a) - $k * [math]::Sin($a)); $p1y = $oy + $hR * ([math]::Sin($a) + $k * [math]::Cos($a))
            $p2x = $ox + $wR * ([math]::Cos($b) + $k * [math]::Sin($b)); $p2y = $oy + $hR * ([math]::Sin($b) - $k * [math]::Cos($b))
            $p3x = $ox + $wR * [math]::Cos($b); $p3y = $oy + $hR * [math]::Sin($b)
            [void]$sb.Append('C' + (& $norm $p1x $p1y) + ' ' + (& $norm $p2x $p2y) + ' ' + (& $norm $p3x $p3y) + ' ')
          }
          $cx = $ox + $wR * [math]::Cos($t1 + $dt); $cy = $oy + $hR * [math]::Sin($t1 + $dt) }
        'close' { [void]$sb.Append('Z ') }
      }
    }
  }
  return $sb.ToString().Trim()
}

$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($src, -1, 0, 0)
  $W = [double]$pres.PageSetup.SlideWidth; $H = [double]$pres.PageSetup.SlideHeight
  $script:media = 0; $script:geoms = @()
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
      if ($type -eq 9) { $o.kind = 'line' }
      $visible = $o.style.fill -or $o.style.gradient -or $o.style.line -or $o.style.complexFill -or $o.style.fx
      $needSnap = $visible -and ($o.style.complexFill -or ($o.style.fx -and $o.kind -ne 'text')) -and $o.kind -ne 'line'
      # outlines the editor has no preset for (block arc, teardrop, freeform ...) are read as a path instead
      $isGeom = ($type -eq 5) -or (($type -eq 1 -or $type -eq 14 -or $type -eq 17) -and $o.style.ast -and -not ($KNOWN_AST -contains $o.style.ast))
      if (-not $needSnap -and $visible -and $isGeom -and ($o.kind -eq 'other' -or $o.kind -eq 'text')) { $script:geoms += [pscustomobject]@{ o = $o; sh = $sh; slide = $script:slideNo } }
      if ($needSnap -and $o.kind -ne 'picture' -and $o.kind -ne 'table' -and $o.kind -ne 'chart') {
        $script:media++
        $sn = 'snap-{0:D3}.png' -f $script:media
        $script:snaps += [pscustomobject]@{ o = $o; file = $sn }   # rendered after the walk (SnapShapes)
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
    $script:snaps = @(); $script:slideNo = $i
    Walk $sl.Shapes '' $list $false
    SnapShapes $sl $script:snaps $W $H
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
  # ---------- outlines of shapes the editor has no preset for ----------
  # Freeforms already carry their path in the file. Other presets (block arc, teardrop ...) are rebuilt at the slide
  # origin and merged with a copy of themselves, which turns them into a freeform with the same outline.
  if ($script:geoms.Count) {
    foreach ($g in $script:geoms) {
      try {
        $ast = [int]$g.o.style.ast
        if ([int]$g.sh.Type -eq 5 -or $ast -eq 138 -or $ast -eq -2) { $g | Add-Member gid ([int]$g.sh.Id); continue }
        $sl = $pres.Slides.Item($g.slide); $fw = [double]$g.sh.Width; $fh = [double]$g.sh.Height
        $t1 = $sl.Shapes.AddShape($ast, 0, 0, [single]$fw, [single]$fh)
        try { for ($a = 1; $a -le $g.sh.Adjustments.Count; $a++) { $t1.Adjustments.Item($a) = $g.sh.Adjustments.Item($a) } } catch {}
        $t2 = $t1.Duplicate().Item(1); $t2.Left = 0; $t2.Top = 0
        $n1 = '__geo_' + [guid]::NewGuid().ToString('N'); $n2 = $n1 + '_b'; $t1.Name = $n1; $t2.Name = $n2
        $before = @{}; foreach ($s in $sl.Shapes) { $before[[int]$s.Id] = 1 }
        $sl.Shapes.Range([object[]]@($n1, $n2)).MergeShapes(1, $t1)   # PowerShell needs the primary shape passed explicitly
        foreach ($s in $sl.Shapes) { if (-not $before.ContainsKey([int]$s.Id)) { $g | Add-Member gid ([int]$s.Id); $g | Add-Member gbox @([double]$s.Left, [double]$s.Top, [double]$s.Width, [double]$s.Height); $g | Add-Member fw $fw; $g | Add-Member fh $fh } }
      } catch {}
    }
    # read the paths from a saved copy (the original file is never changed)
    $copy = Join-Path $tmp 'geom.pptx'
    try {
      $pres.SaveCopyAs($copy)
      Add-Type -AssemblyName System.IO.Compression.FileSystem
      $zip = [IO.Compression.ZipFile]::OpenRead($copy)
      try {
        $readXml = { param($name) $en = $zip.GetEntry($name); if (-not $en) { return $null }; $sr = New-Object IO.StreamReader($en.Open()); $x = New-Object Xml.XmlDocument; $x.LoadXml($sr.ReadToEnd()); $sr.Close(); return $x }
        $presX = & $readXml 'ppt/presentation.xml'; $relX = & $readXml 'ppt/_rels/presentation.xml.rels'
        $ns = New-Object Xml.XmlNamespaceManager($presX.NameTable)
        $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main'); $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'); $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $targets = @{}; foreach ($rel in $relX.DocumentElement.ChildNodes) { $targets[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target') }
        $slideFiles = @(); foreach ($sid in $presX.SelectNodes('//p:sldIdLst/p:sldId', $ns)) { $slideFiles += ('ppt/' + $targets[$sid.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')]) }
        $docs = @{}
        foreach ($g in $script:geoms) {
          if (-not $g.gid) { continue }
          try {
            if (-not $docs.ContainsKey($g.slide)) { $docs[$g.slide] = & $readXml $slideFiles[$g.slide - 1] }
            $d = $docs[$g.slide]; $nsl = New-Object Xml.XmlNamespaceManager($d.NameTable)
            foreach ($pfx in 'p', 'a') { $nsl.AddNamespace($pfx, $ns.LookupNamespace($pfx)) }
            $sp = $d.SelectSingleNode("//p:sp[p:nvSpPr/p:cNvPr/@id='$($g.gid)']", $nsl)
            $pl = $sp.SelectSingleNode('p:spPr/a:custGeom/a:pathLst', $nsl); if (-not $pl) { continue }
            $ext = $sp.SelectSingleNode('p:spPr/a:xfrm/a:ext', $nsl)
            $geo = GeomFromPathLst $pl $nsl ([double]$ext.GetAttribute('cx')) ([double]$ext.GetAttribute('cy')) $g.gbox $g.fw $g.fh
            if ($geo) { $g.o.geom = $geo }
          } catch {}
        }
      } finally { $zip.Dispose() }
    } catch {}
    # anything whose outline could not be read is kept as an exact cut-out picture instead
    $bySlide = @{}
    foreach ($g in $script:geoms) { if (-not $g.o.geom) { $script:media++; if (-not $bySlide.ContainsKey($g.slide)) { $bySlide[$g.slide] = @() }; $bySlide[$g.slide] += [pscustomobject]@{ o = $g.o; file = ('snap-{0:D3}.png' -f $script:media) } } }
    foreach ($k in $bySlide.Keys) { SnapShapes $pres.Slides.Item($k) $bySlide[$k] $W $H }
  }
  $info = [ordered]@{ width = $W; height = $H; slideCount = $pres.Slides.Count; slides = $slides }
  $pres.Saved = -1; $pres.Close()
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
