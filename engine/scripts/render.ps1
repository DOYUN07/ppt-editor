param(
  [Parameter(Mandatory=$true)][string]$Doc,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$PngDir,
  [switch]$Potx,
  [switch]$Pdf
)
# Draws a fully laid-out document (every element with its position, style and theme-color tokens) into PowerPoint.
# The editor computes all geometry; this script only translates elements into native PowerPoint objects so the
# result stays editable: theme colors/fonts, a title placeholder per slide, native tables/charts, transitions, animations.
#
# Doc JSON: { size:{w,h}, theme:{name, colors{dk1..folHlink}, headFont, bodyFont}, slides:[ { elements:[...], bg:{color|image},
#   transition:{effect,dur}, notes, hidden, section } ] }
# Element: { kind:text|shape|line|image|icon|table|chart, x,y,w,h, rot, flipH, flipV, opacity, isTitle,
#   shape, radius, fill, line:{color,width,dash,start,end}, shadow, text:{paras:[{runs:[{t,b,i,u,s,color,size,font,link}],align,level,bullet}],
#   size,color,font,bold,italic,align,valign,lineSpacing,spaceAfter,pad,field}, image:{src,fit,frame}, icon:{svg}, table:{...}, chart:{...}, anim:{...} }
# Colors are theme tokens (dk1, lt1, dk2, lt2, accent1..accent6) or '#RRGGBB'.
$ErrorActionPreference = 'Stop'
function LongPath([string]$p) { $full = [IO.Path]::GetFullPath($p); if ($full.Length -gt 240 -and -not $full.StartsWith('\\?\')) { return '\\?\' + $full }; return $full }
function Has($o, [string]$name) { return ($null -ne $o) -and ($o.PSObject.Properties.Name -contains $name) -and ($null -ne $o.$name) -and ("$($o.$name)" -ne '') }
function Get($o, [string]$name, $default) { if (Has $o $name) { return $o.$name }; return $default }
function IsTrue($v) { return ([string]$v) -eq 'True' }

$docObj = [IO.File]::ReadAllText((LongPath $Doc), [Text.Encoding]::UTF8) | ConvertFrom-Json
$W = [double](Get $docObj.size 'w' 960); $H = [double](Get $docObj.size 'h' 540)
$theme = $docObj.theme

$TOKENS = @{ dk1 = 13; lt1 = 14; dk2 = 15; lt2 = 16; accent1 = 5; accent2 = 6; accent3 = 7; accent4 = 8; accent5 = 9; accent6 = 10; tx1 = 13; bg1 = 14; tx2 = 15; bg2 = 16 }
function RGBc([string]$hex) { $hex = $hex.TrimStart('#'); return [int]([Convert]::ToInt32($hex.Substring(0,2),16) + [Convert]::ToInt32($hex.Substring(2,2),16)*256 + [Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function SetColor($colorFormat, [string]$c) {
  if (-not $c) { return }
  if ($TOKENS.ContainsKey($c)) { $colorFormat.ObjectThemeColor = $TOKENS[$c] } else { $colorFormat.RGB = (RGBc $c) }
}
function HexOf([string]$c) { if ($TOKENS.ContainsKey($c)) { $k = $c; if ($k -eq 'tx1') { $k = 'dk1' }; return [string]$theme.colors.$k } ; return $c.TrimStart('#') }

$SHAPES = @{ rect = 1; roundRect = 5; ellipse = 9; triangle = 7; rtTriangle = 8; diamond = 4; parallelogram = 2; trapezoid = 3; pentagon = 12; hexagon = 10; octagon = 6;
  star4 = 91; star5 = 92; star6 = 147; plus = 11; heart = 21; arrowRight = 33; arrowLeft = 34; arrowUp = 35; arrowDown = 36; arrowLR = 37; chevron = 52; homePlate = 51;
  callout = 106; cloud = 179; donut = 18; lightning = 22; frame = 158; snipRect = 155; round2Same = 152; can = 13; cube = 14; moon = 24; sun = 23; wave = 103; plaque = 28 }
$TRANS = @{ fade = 3849; fadeBlack = 1793; push = 3853; pushUp = 3855; wipe = 2817; split = 3585; reveal = 3894; cover = 1281; uncover = 2049; zoom = 3345; randomBars = 2305;
  dissolve = 1537; flip = 3905; cube = 3914; gallery = 3880; curtains = 3938; vortex = 3863; ripple = 3867; honeycomb = 3898; doors = 3884; window = 3886; ferris = 3899;
  conveyor = 3882; rotate = 3918; orbit = 3926; fly = 3890; switch = 3901; pan = 3930; wind = 3939; prestige = 3941; fracture = 3942; crush = 3943; peel = 3944; pageCurl = 3946; airplane = 3950; origami = 3952; random = 513 }
$ANIMS = @{ appear = 1; fade = 10; fly = 2; wipe = 22; zoom = 23; float = 30; split = 16; wheel = 21; bounce = 26; growTurn = 31; riseUp = 34; swivel = 19; peek = 12;
  dissolve = 9; randomBars = 14; fadedZoom = 48; expand = 50; spin = 61; growShrink = 59; teeter = 80; flash = 11 }
$DIRS = @{ bottom = 11; left = 4; right = 2; top = 10 }
$CHARTS = @{ column = 51; columnStacked = 52; bar = 57; barStacked = 58; line = 65; lineSmooth = 4; pie = 5; doughnut = -4120; area = 1; areaStacked = 76; radar = -4151 }
$DASH = @{ solid = 1; dash = 4; dot = 3; dashDot = 5; longDash = 7 }

$EngineDir = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP ('ppt-render-' + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Force $tmp | Out-Null
$script:staged = @{}
# PowerPoint is a separate process and reliably reads %TEMP%; copy every image there first.
function Stage([string]$p) {
  if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $EngineDir $p }
  $dst = Join-Path $tmp ('img-' + $script:staged.Count + [IO.Path]::GetExtension($p))
  Copy-Item -LiteralPath (LongPath $p) -Destination $dst -Force
  $script:staged[$p] = $dst; return $dst
}
function WriteSvg([string]$svg) { $f = Join-Path $tmp ('icon-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.svg'); [IO.File]::WriteAllText($f, $svg, (New-Object Text.UTF8Encoding $false)); return $f }

function ApplyFill($sh, $e) {
  if (Has $e 'fill') {
    $sh.Fill.Visible = -1; $sh.Fill.Solid(); SetColor $sh.Fill.ForeColor ([string]$e.fill)
    $op = [double](Get $e 'fillOpacity' 1); if ($op -lt 1) { $sh.Fill.Transparency = [single](1 - $op) }
  } else { $sh.Fill.Visible = 0 }
}
function ApplyLine($sh, $ln) {
  if ($ln -and (Has $ln 'color') -and [double](Get $ln 'width' 1) -gt 0) {
    $sh.Line.Visible = -1; SetColor $sh.Line.ForeColor ([string]$ln.color); $sh.Line.Weight = [single](Get $ln 'width' 1)
    $d = [string](Get $ln 'dash' 'solid'); if ($DASH.ContainsKey($d)) { $sh.Line.DashStyle = $DASH[$d] }
  } else { $sh.Line.Visible = 0 }
}
function ApplyCommon($sh, $e) {
  if (Has $e 'rot') { $sh.Rotation = [single]$e.rot }
  if (IsTrue (Get $e 'flipH' $false)) { $sh.Flip(0) }
  if (IsTrue (Get $e 'flipV' $false)) { $sh.Flip(1) }
  if (IsTrue (Get $e 'shadow' $false)) {
    $sh.Shadow.Visible = -1; $sh.Shadow.ForeColor.RGB = 0; $sh.Shadow.OffsetX = 0; $sh.Shadow.OffsetY = 4; $sh.Shadow.Blur = 12; $sh.Shadow.Transparency = 0.75
  } else { try { $sh.Shadow.Visible = 0 } catch {} }
  if (Has $e 'name') { try { $sh.Name = [string]$e.name } catch {} }
}
function FontFor([string]$f) { if ($f -eq 'head') { return '+mj-lt' } ; if ($f -eq 'body' -or -not $f) { return '+mn-lt' }; return $f }
function SetFont($font, [string]$f) {
  $n = FontFor $f
  $font.Name = $n
  if ($n -eq '+mj-lt') { $font.NameFarEast = '+mj-ea' } elseif ($n -eq '+mn-lt') { $font.NameFarEast = '+mn-ea' } else { $font.NameFarEast = $n }
}
$ALIGN = @{ left = 1; center = 2; right = 3; justify = 4 }
$VALIGN = @{ top = 1; middle = 3; bottom = 4 }
function FillText($sh, $tx, [bool]$isShape) {
  $tf = $sh.TextFrame
  $tf.WordWrap = $(if ((Get $tx 'wrap' $true) -eq $false) { 0 } else { -1 }); $tf.AutoSize = 0
  $pad = Get $tx 'pad' $null
  if ($pad) { $p4 = @($pad); $tf.MarginTop = [single]$p4[0]; $tf.MarginRight = [single]$p4[1]; $tf.MarginBottom = [single]$p4[2]; $tf.MarginLeft = [single]$p4[3] }
  elseif (-not $isShape) { $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0 }
  $tf.VerticalAnchor = $VALIGN[[string](Get $tx 'valign' 'top')]
  $paras = @($tx.paras)
  $strings = @()
  foreach ($p in $paras) { $strings += ((@($p.runs) | ForEach-Object { ([string]$_.t).Replace("`n", [string][char]11) }) -join '') }
  $tr = $tf.TextRange
  $tr.Text = ($strings -join "`r")
  # element defaults
  $tr.Font.Size = [single](Get $tx 'size' 16)
  SetFont $tr.Font ([string](Get $tx 'font' 'body'))
  SetColor $tr.Font.Color ([string](Get $tx 'color' 'dk1'))
  $tr.Font.Bold = [int](IsTrue (Get $tx 'bold' $false)) * -1
  $tr.Font.Italic = [int](IsTrue (Get $tx 'italic' $false)) * -1
  $ls = [double](Get $tx 'lineSpacing' 1.1)
  for ($i = 1; $i -le $paras.Count; $i++) {
    $p = $paras[$i-1]; $pr = $tr.Paragraphs($i)
    $pr.ParagraphFormat.Alignment = $ALIGN[[string](Get $p 'align' (Get $tx 'align' 'left'))]
    $pr.ParagraphFormat.SpaceWithin = [single]$ls
    $pr.ParagraphFormat.WordWrap = -1
    $sa = Get $p 'spaceAfter' (Get $tx 'spaceAfter' $null); if ($null -ne $sa) { $pr.ParagraphFormat.LineRuleAfter = 0; $pr.ParagraphFormat.SpaceAfter = [single]$sa }
    $sb = Get $p 'spaceBefore' $null; if ($null -ne $sb -and $i -gt 1) { $pr.ParagraphFormat.LineRuleBefore = 0; $pr.ParagraphFormat.SpaceBefore = [single]$sb }
    $lvl = [int](Get $p 'level' 1); if ($lvl -gt 1) { $pr.IndentLevel = $lvl }
    $bul = [string](Get $p 'bullet' 'none')
    if ($bul -eq 'disc' -or $bul -eq 'square' -or $bul -eq 'char') {
      $pr.ParagraphFormat.Bullet.Visible = -1; $pr.ParagraphFormat.Bullet.Type = 1
      if ($bul -eq 'char') {
        $pr.ParagraphFormat.Bullet.Font.Name = 'Segoe UI Symbol'
        $pr.ParagraphFormat.Bullet.Character = [int][char]([string](Get $p 'bulletChar' ([string][char]8226)))[0]
        $pr.ParagraphFormat.Bullet.RelativeSize = 1
      } else {
        $pr.ParagraphFormat.Bullet.Character = $(if ($bul -eq 'square') { 9632 } else { 9679 })
        $pr.ParagraphFormat.Bullet.RelativeSize = 0.7
      }
      if (Has $tx 'bulletColor') { $pr.ParagraphFormat.Bullet.UseTextColor = 0; SetColor $pr.ParagraphFormat.Bullet.Font.Color ([string]$tx.bulletColor) } else { $pr.ParagraphFormat.Bullet.UseTextColor = -1 }
    } elseif ($bul -eq 'num') {
      $pr.ParagraphFormat.Bullet.Visible = -1; $pr.ParagraphFormat.Bullet.Type = 2; $pr.ParagraphFormat.Bullet.Style = 3
    } else { $pr.ParagraphFormat.Bullet.Visible = 0 }
    if ($bul -ne 'none') { $sh.TextFrame.Ruler.Levels($lvl).FirstMargin = [single](($lvl - 1) * 22); $sh.TextFrame.Ruler.Levels($lvl).LeftMargin = [single](($lvl - 1) * 22 + 16) }
    # runs
    $pos = 1
    foreach ($r in @($p.runs)) {
      $len = ([string]$r.t).Length
      if ($len -gt 0) {
        $c = $pr.Characters($pos, $len)
        if (Has $r 'size') { $c.Font.Size = [single]$r.size }
        if (Has $r 'font') { SetFont $c.Font ([string]$r.font) }
        if (Has $r 'color') { SetColor $c.Font.Color ([string]$r.color) }
        if (Has $r 'b') { $c.Font.Bold = [int](IsTrue $r.b) * -1 }
        if (Has $r 'i') { $c.Font.Italic = [int](IsTrue $r.i) * -1 }
        if (Has $r 'u') { $c.Font.Underline = [int](IsTrue $r.u) * -1 }
        if ((Has $r 's') -and (IsTrue $r.s)) { try { $sh.TextFrame2.TextRange.Paragraphs($i).Characters($pos, $len).Font.Strike = 1 } catch {} }
        if (Has $r 'link') { try { $c.ActionSettings(1).Hyperlink.Address = [string]$r.link } catch {} }
      }
      $pos += $len
    }
  }
  if ((Get $tx 'field' '') -eq 'page') { $tr.Text = ''; $null = $tr.InsertSlideNumber(); $tr.Font.Size = [single](Get $tx 'size' 9); SetColor $tr.Font.Color ([string](Get $tx 'color' 'accent6')); $tr.ParagraphFormat.Alignment = $ALIGN[[string](Get $tx 'align' 'right')] }
}

function PlacePicture($slide, [string]$path, $e) {
  $l = [double]$e.x; $t = [double]$e.y; $w = [double]$e.w; $h = [double]$e.h
  $pic = $slide.Shapes.AddPicture((Stage $path), 0, -1, [single]$l, [single]$t)
  $pic.LockAspectRatio = 0
  $fit = [string](Get $e.image 'fit' 'cover')
  $w0 = [double]$pic.Width; $h0 = [double]$pic.Height
  if ($fit -eq 'contain') {
    $r = [math]::Min($w / $w0, $h / $h0); $nw = $w0 * $r; $nh = $h0 * $r
    $pic.Width = [single]$nw; $pic.Height = [single]$nh; $pic.Left = [single]($l + ($w - $nw) / 2); $pic.Top = [single]($t + ($h - $nh) / 2)
  } elseif ($fit -eq 'stretch') {
    $pic.Left = [single]$l; $pic.Top = [single]$t; $pic.Width = [single]$w; $pic.Height = [single]$h
  } else {
    $r = [math]::Max($w / $w0, $h / $h0)
    $fx = [double](Get $e.image 'focusX' 0.5); $fy = [double](Get $e.image 'focusY' 0.5)
    $cw = $w0 - $w / $r; $chh = $h0 - $h / $r
    $pic.PictureFormat.CropLeft = [single]($cw * $fx); $pic.PictureFormat.CropRight = [single]($cw * (1 - $fx))
    $pic.PictureFormat.CropTop = [single]($chh * $fy); $pic.PictureFormat.CropBottom = [single]($chh * (1 - $fy))
    $pic.Left = [single]$l; $pic.Top = [single]$t; $pic.Width = [single]$w; $pic.Height = [single]$h
  }
  $frame = [string](Get $e 'shape' 'rect')
  if ($frame -ne 'rect' -and $SHAPES.ContainsKey($frame)) { $pic.AutoShapeType = $SHAPES[$frame]; if ($frame -eq 'roundRect') { $pic.Adjustments.Item(1) = [single](Get $e 'radius' 0.1) } }
  ApplyLine $pic (Get $e 'line' $null)
  return $pic
}

function DrawTable($slide, $e) {
  $tb = $e.table
  $rows = @($tb.rows); $nr = $rows.Count; $nc = (@($rows[0])).Count
  $shp = $slide.Shapes.AddTable($nr, $nc, [single]$e.x, [single]$e.y, [single]$e.w, [single]$e.h)
  $t = $shp.Table
  $t.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}')
  $cw = @(Get $tb 'colW' @())
  for ($c = 1; $c -le $nc; $c++) { if ($cw.Count -ge $c) { $t.Columns.Item($c).Width = [single]$cw[$c-1] } else { $t.Columns.Item($c).Width = [single]([double]$e.w / $nc) } }
  $rh = [double]$e.h / $nr
  $header = IsTrue (Get $tb 'header' $true)
  $fs = [single](Get $tb 'fontSize' 14)
  $status = Get $tb 'statusColors' $null
  for ($r = 1; $r -le $nr; $r++) {
    $t.Rows.Item($r).Height = [single]$rh
    for ($c = 1; $c -le $nc; $c++) {
      $cell = $t.Cell($r, $c); $tf = $cell.Shape.TextFrame
      $tf.VerticalAnchor = 3; $tf.MarginLeft = 12; $tf.MarginRight = 12
      $tr = $tf.TextRange; $tr.Text = [string](@($rows[$r-1])[$c-1]); $tr.Font.Size = $fs
      SetFont $tr.Font 'body'
      $tr.ParagraphFormat.Alignment = $ALIGN[[string](Get $tb 'align' 'left')]
      $cell.Shape.Fill.Visible = -1; $cell.Shape.Fill.Solid()
      if ($header -and $r -eq 1) {
        SetColor $cell.Shape.Fill.ForeColor ([string](Get $tb 'headerFill' 'dk2'))
        SetColor $tr.Font.Color ([string](Get $tb 'headerColor' 'lt1')); $tr.Font.Bold = -1
      } else {
        $band = ($r % 2 -eq 1) -and (IsTrue (Get $tb 'banded' $true))
        if ($band) { SetColor $cell.Shape.Fill.ForeColor ([string](Get $tb 'bandFill' 'lt2')) } else { SetColor $cell.Shape.Fill.ForeColor ([string](Get $tb 'bodyFill' 'lt1')) }
        SetColor $tr.Font.Color ([string](Get $tb 'color' 'dk1'))
        if ($c -eq 1 -and (IsTrue (Get $tb 'firstColBold' $true))) { $tr.Font.Bold = -1 }
        if ($status -and ($status.PSObject.Properties.Name -contains $tr.Text)) { SetColor $tr.Font.Color ([string]$status.($tr.Text)); $tr.Font.Bold = -1 }
      }
      if (Has $tb 'borderColor') { $b = $cell.Borders.Item(3); $b.Visible = -1; $b.Weight = 0.75; SetColor $b.ForeColor ([string]$tb.borderColor) }
    }
  }
  return $shp
}

function DrawChart($slide, $e) {
  $ch = $e.chart
  $kind = [string](Get $ch 'type' 'column'); $ct = $CHARTS[$kind]; if (-not $ct) { $ct = 51 }; if ($kind -eq 'doughnut') { $ct = 5 }
  # Excel starts lazily; retry the chart insert and wait for its data workbook.
  $chs = $null
  for ($try = 0; $try -lt 3 -and -not $chs; $try++) { try { $chs = $slide.Shapes.AddChart2(-1, $ct, [single]$e.x, [single]$e.y, [single]$e.w, [single]$e.h) } catch { Start-Sleep -Seconds 2 } }
  if (-not $chs) { throw 'chart could not be inserted' }
  $c = $chs.Chart
  if ($kind -eq 'doughnut') { try { $c.ChartType = -4120 } catch {} }   # AddChart2 rejects xlDoughnut; insert as pie and switch
  $wb = $null
  for ($try = 0; $try -lt 20 -and -not $wb; $try++) { try { $c.ChartData.Activate(); $wb = $c.ChartData.Workbook } catch {}; if (-not $wb) { Start-Sleep -Milliseconds 700 } }
  if (-not $wb) { throw 'chart data workbook did not open' }
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('A1:Z60').Clear()
  $cats = @($ch.categories); $series = @($ch.series)
  $ws.Cells.Item(1, 1).Value2 = ' '
  for ($s = 0; $s -lt $series.Count; $s++) { $ws.Cells.Item(1, $s + 2).Value2 = [string](Get $series[$s] 'name' "Series $($s+1)") }
  for ($i = 0; $i -lt $cats.Count; $i++) {
    # categories are text: '2020' must not become a number (Excel would turn the column into a data series)
    $ws.Cells.Item($i + 2, 1).NumberFormat = '@'; $ws.Cells.Item($i + 2, 1).Value2 = [string]$cats[$i]
    for ($s = 0; $s -lt $series.Count; $s++) { $v = @($series[$s].values); if ($i -lt $v.Count) { $ws.Cells.Item($i + 2, $s + 2).Value2 = [double]$v[$i] } }
  }
  $lastCol = [char](65 + $series.Count); $lastRow = $cats.Count + 1
  try { $null = $ws.ListObjects.Item(1).Resize($ws.Range("A1:$lastCol$lastRow")) } catch {}
  $c.SetSourceData("='" + $ws.Name + "'!`$A`$1:`$$lastCol`$$lastRow", 2)
  $wb.Close(); Start-Sleep -Seconds 3
  $c.HasTitle = $true; $c.ChartTitle.Delete(); $c.HasTitle = $false
  if (Has $ch 'title') { $c.HasTitle = $true; $c.ChartTitle.Text = [string]$ch.title; $c.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14 }
  $c.ChartArea.Format.Line.Visible = 0; $c.ChartArea.Format.Fill.Visible = 0
  $isPie = ($kind -eq 'pie' -or $kind -eq 'doughnut')
  $c.HasLegend = ($isPie -or $series.Count -gt 1) -and -not ((Has $ch 'legend') -and -not (IsTrue $ch.legend))
  if ($c.HasLegend) { $c.Legend.Position = -4107; $c.Legend.Format.TextFrame2.TextRange.Font.Size = 11 }
  $palette = @(Get $ch 'colors' @('dk2','accent1','accent2','accent3','accent4','accent6'))
  for ($s = 1; $s -le $series.Count; $s++) {
    $ser = $c.SeriesCollection($s)
    $col = HexOf ([string](Get $series[$s-1] 'color' $palette[($s-1) % $palette.Count]))
    if ($isPie) {
      for ($i = 1; $i -le $cats.Count; $i++) { $ser.Points($i).Format.Fill.ForeColor.RGB = (RGBc (HexOf ([string]$palette[($i-1) % $palette.Count]))) }
    } elseif ($kind -like 'line*' -or $kind -eq 'radar') {
      $ser.Format.Line.ForeColor.RGB = (RGBc $col); $ser.Format.Line.Weight = 3
      try { $ser.MarkerStyle = 8; $ser.MarkerSize = 7; $ser.MarkerBackgroundColor = (RGBc $col); $ser.MarkerForegroundColor = (RGBc $col) } catch {}
    } else {
      $ser.Format.Fill.Solid(); $ser.Format.Fill.ForeColor.RGB = (RGBc $col)
      if ($s -eq 1 -and (Has $ch 'pointColors')) { $pcs = @($ch.pointColors); for ($i = 1; $i -le [math]::Min($cats.Count, $pcs.Count); $i++) { try { $ser.Points($i).Format.Fill.Solid(); $ser.Points($i).Format.Fill.ForeColor.RGB = (RGBc (HexOf ([string]$pcs[$i-1]))) } catch {} } }
      if ($s -eq 1 -and (Has $ch 'highlight') -and [int]$ch.highlight -ge 0 -and [int]$ch.highlight -lt $cats.Count) { $ser.Points([int]$ch.highlight + 1).Format.Fill.ForeColor.RGB = (RGBc (HexOf 'accent1')) }
    }
    if (IsTrue (Get $ch 'showValues' $true)) {
      $ser.HasDataLabels = $true; $dl = $ser.DataLabels()
      $dl.Format.TextFrame2.TextRange.Font.Size = 11; $dl.Format.TextFrame2.TextRange.Font.Bold = -1
      if (Has $ch 'numFmt') { try { $dl.NumberFormat = [string]$ch.numFmt } catch {} }
      try { $dl.Format.TextFrame2.WordWrap = 0 } catch {}
      if ($isPie) { $dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (RGBc 'FFFFFF') } else { $dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (RGBc (HexOf 'dk1')) }
    }
  }
  if (-not $isPie) {
    try { $c.ChartGroups(1).GapWidth = 70 } catch {}
    try {
      $va = $c.Axes(2); $va.MajorGridlines.Format.Line.ForeColor.RGB = (RGBc 'E3E7EE'); $va.Format.Line.Visible = 0; $va.TickLabels.Font.Size = 10; $va.TickLabels.Font.Color = (RGBc (HexOf 'accent6'))
      $ca = $c.Axes(1); $ca.TickLabels.Font.Size = 11; $ca.TickLabels.Font.Color = (RGBc (HexOf 'dk1')); $ca.Format.Line.ForeColor.RGB = (RGBc 'C9D0DB')
    } catch {}
    try { if (Has $ch 'numFmt') { $c.Axes(2).TickLabels.NumberFormat = [string]$ch.numFmt } } catch {}
    try { if (Has $ch 'axisMin') { $c.Axes(2).MinimumScale = [double]$ch.axisMin }; if (Has $ch 'axisMax') { $c.Axes(2).MaximumScale = [double]$ch.axisMax } } catch {}
    try { if ((Get $ch 'gridlines' $true) -eq $false) { $c.Axes(2).HasMajorGridlines = $false } } catch {}
    try { if ((Get $ch 'valueAxis' $true) -eq $false) { $c.Axes(2).TickLabelPosition = -4142 } } catch {}
  }
  return $chs
}

function AddAnim($slide, $sh, $a) {
  $eff = [string](Get $a 'effect' 'none'); if ($eff -eq 'none' -or -not $ANIMS.ContainsKey($eff)) { return }
  $trig = @{ click = 1; with = 2; after = 3 }[[string](Get $a 'trigger' 'click')]; if (-not $trig) { $trig = 1 }
  $fx = $slide.TimeLine.MainSequence.AddEffect($sh, $ANIMS[$eff], 0, $trig)
  if ($eff -eq 'fly' -or $eff -eq 'wipe' -or $eff -eq 'peek') { $d = [string](Get $a 'dir' 'bottom'); if ($DIRS.ContainsKey($d)) { try { $fx.EffectParameters.Direction = $DIRS[$d] } catch {} } }
  $fx.Timing.Duration = [single](Get $a 'dur' 0.5)
  $fx.Timing.TriggerDelayTime = [single](Get $a 'delay' 0)
}

$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Add(-1)
  try { if ($app.Presentations.Count -le 1) { $app.WindowState = 2 } } catch {}
  $pres.PageSetup.SlideWidth = [single]$W; $pres.PageSetup.SlideHeight = [single]$H
  $master = $pres.SlideMaster
  $master.Design.Name = [string](Get $theme 'name' 'Custom')
  $order = @('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
  for ($i = 0; $i -lt 12; $i++) { $master.Theme.ThemeColorScheme.Colors($i+1).RGB = (RGBc ([string]$theme.colors.($order[$i]))) }
  $fs = $master.Theme.ThemeFontScheme
  $fs.MajorFont.Item(1).Name = [string]$theme.headFont; $fs.MajorFont.Item(3).Name = [string]$theme.headFont
  $fs.MinorFont.Item(1).Name = [string]$theme.bodyFont; $fs.MinorFont.Item(3).Name = [string]$theme.bodyFont
  $master.Background.Fill.Solid(); $master.Background.Fill.ForeColor.ObjectThemeColor = 14
  # Footer, date and slide-number placeholders are drawn as elements by the editor; hide the master ones.
  foreach ($sh in @($master.Shapes)) { if ($sh.Type -eq 14 -and @(13, 15, 16) -contains $sh.PlaceholderFormat.Type) { $sh.Delete() } }
  $LT = $master.CustomLayouts.Item(6)   # Title Only: keeps a real title on every slide (outline, accessibility)
  $LB = $master.CustomLayouts.Item(7)   # Blank
  foreach ($lay in @($LT, $LB)) { foreach ($sh in @($lay.Shapes)) { if ($sh.Type -eq 14 -and @(13, 15, 16) -contains $sh.PlaceholderFormat.Type) { $sh.Delete() } } }

  $pending = @()
  $sections = @()
  $n = 0
  foreach ($sd in @($docObj.slides)) {
    $n++
    $els = @($sd.elements)
    $titleEl = $els | Where-Object { IsTrue (Get $_ 'isTitle' $false) } | Select-Object -First 1
    $lay = $LB; if ($titleEl) { $lay = $LT }
    $slide = $pres.Slides.AddSlide($n, $lay)
    if (Has $sd 'bg') {
      $slide.FollowMasterBackground = 0
      if (Has $sd.bg 'image') { $slide.Background.Fill.UserPicture((Stage ([string]$sd.bg.image))) }
      elseif (Has $sd.bg 'color') { $slide.Background.Fill.Solid(); SetColor $slide.Background.Fill.ForeColor ([string]$sd.bg.color) }
    }
    $titlePh = $null
    if ($titleEl) { foreach ($sh in @($slide.Shapes)) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 1) { $titlePh = $sh } } }
    $animList = @()
    foreach ($e in $els) {
      $kind = [string]$e.kind; $sh = $null
      try {
        switch ($kind) {
          'text' {
            if ($titlePh -and $e -eq $titleEl) { $sh = $titlePh; $sh.Left = [single]$e.x; $sh.Top = [single]$e.y; $sh.Width = [single]$e.w; $sh.Height = [single]$e.h; $sh.ZOrder(0) }
            else { $sh = $slide.Shapes.AddTextbox(1, [single]$e.x, [single]$e.y, [single]$e.w, [single]$e.h) }
            ApplyFill $sh $e; ApplyLine $sh (Get $e 'line' $null)
            FillText $sh $e.text $false
            $sh.Left = [single]$e.x; $sh.Top = [single]$e.y; $sh.Width = [single]$e.w; $sh.Height = [single]$e.h
          }
          'shape' {
            $type = $SHAPES[[string](Get $e 'shape' 'rect')]; if (-not $type) { $type = 1 }
            $sh = $slide.Shapes.AddShape($type, [single]$e.x, [single]$e.y, [single]$e.w, [single]$e.h)
            if ($type -eq 5 -or $type -eq 155 -or $type -eq 152) { $sh.Adjustments.Item(1) = [single](Get $e 'radius' 0.12) }
            ApplyFill $sh $e; ApplyLine $sh (Get $e 'line' $null)
            if ((Has $e 'text') -and @($e.text.paras).Count -gt 0) { FillText $sh $e.text $true }
          }
          'line' {
            $x1 = [double]$e.x; $y1 = [double]$e.y; $x2 = $x1 + [double]$e.w; $y2 = $y1 + [double]$e.h
            if (IsTrue (Get $e 'flipH' $false)) { $t = $x1; $x1 = $x2; $x2 = $t }
            if (IsTrue (Get $e 'flipV' $false)) { $t = $y1; $y1 = $y2; $y2 = $t }
            $sh = $slide.Shapes.AddLine([single]$x1, [single]$y1, [single]$x2, [single]$y2)
            $ln = Get $e 'line' $null
            SetColor $sh.Line.ForeColor ([string](Get $ln 'color' 'dk1')); $sh.Line.Weight = [single](Get $ln 'width' 2)
            $d = [string](Get $ln 'dash' 'solid'); if ($DASH.ContainsKey($d)) { $sh.Line.DashStyle = $DASH[$d] }
            if ((Get $ln 'end' 'none') -eq 'arrow') { $sh.Line.EndArrowheadStyle = 2 }
            if ((Get $ln 'start' 'none') -eq 'arrow') { $sh.Line.BeginArrowheadStyle = 2 }
            $e = $e | Select-Object * -ExcludeProperty flipH, flipV
          }
          'image' { $sh = PlacePicture $slide ([string]$e.image.src) $e }
          'icon' { $sh = $slide.Shapes.AddPicture((WriteSvg ([string]$e.icon.svg)), 0, -1, [single]$e.x, [single]$e.y, [single]$e.w, [single]$e.h) }
          'table' { $sh = DrawTable $slide $e }
          'chart' { $pending += ,@($slide, $e) }
        }
        if ($sh) {
          ApplyCommon $sh $e
          if ((Has $e 'opacity') -and [double]$e.opacity -lt 1 -and $kind -ne 'image' -and $kind -ne 'icon') { try { $sh.Fill.Transparency = [single](1 - [double]$e.opacity) } catch {} }
          if (Has $e 'anim') { $animList += [pscustomobject]@{ Shape = $sh; Anim = $e.anim } }
        }
      } catch { Write-Output ("WARN slide $n element ${kind}: " + $_) }
    }
    # Empty title placeholder (title element removed) would show "Click to add title" in edit mode: delete it.
    foreach ($sh in @($slide.Shapes)) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 1 -and -not $sh.TextFrame.HasText) { $sh.Delete() } }
    foreach ($a in @($animList | Sort-Object { [double](Get $_.Anim 'order' 0) })) { try { AddAnim $slide $a.Shape $a.Anim } catch { Write-Output "WARN slide $n animation: $_" } }
    if (Has $sd 'transition') {
      $tEff = [string](Get $sd.transition 'effect' 'none')
      if ($TRANS.ContainsKey($tEff)) { $slide.SlideShowTransition.EntryEffect = $TRANS[$tEff]; $slide.SlideShowTransition.Duration = [single](Get $sd.transition 'dur' 0.7) }
    }
    if (IsTrue (Get $sd 'hidden' $false)) { $slide.SlideShowTransition.Hidden = -1 }
    if (Has $sd 'notes') { $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = ([string]$sd.notes).Replace("`n", "`r") }
    if (Has $sd 'section') { $sections += ,@($n, [string]$sd.section) }
  }

  # Charts last: their Excel session disturbs other COM calls.
  foreach ($pc in $pending) {
    try {
      $sh = DrawChart $pc[0] $pc[1]
      ApplyCommon $sh $pc[1]
      if (Has $pc[1] 'anim') { AddAnim $pc[0] $sh $pc[1].anim }
    } catch { Write-Output ("WARN chart: " + $_) }
    Start-Sleep -Seconds 2
  }
  if ($sections.Count -gt 0) {
    if ($sections[0][0] -ne 1) { $null = $pres.SectionProperties.AddBeforeSlide(1, 'Intro') }
    foreach ($s in $sections) { $null = $pres.SectionProperties.AddBeforeSlide([int]$s[0], [string]$s[1]) }
  }

  $name = [IO.Path]::GetFileNameWithoutExtension($Out)
  $tpptx = Join-Path $tmp ($name + '.pptx')
  $pres.SaveAs($tpptx, 24)
  if ($Potx) { $pres.SaveCopyAs((Join-Path $tmp ($name + '.potx')), 26) }
  if ($Pdf) { $pres.SaveCopyAs((Join-Path $tmp ($name + '.pdf')), 32) }
  if ($PngDir) {
    New-Item -ItemType Directory -Force (LongPath $PngDir) | Out-Null
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
      $png = Join-Path $tmp ("slide-{0:D2}.png" -f $i)
      $pres.Slides.Item($i).Export($png, 'PNG', 1280, [int](1280 * $H / $W))
      Copy-Item -LiteralPath $png -Destination (LongPath (Join-Path $PngDir ("slide-{0:D2}.png" -f $i))) -Force
    }
  }
  $pres.Close()
  $outFull = LongPath $Out
  New-Item -ItemType Directory -Force (Split-Path $outFull -Parent) | Out-Null
  Copy-Item -LiteralPath $tpptx -Destination $outFull -Force
  if ($Potx) { Copy-Item -LiteralPath (Join-Path $tmp ($name + '.potx')) -Destination ([IO.Path]::ChangeExtension($outFull, '.potx')) -Force }
  if ($Pdf) { Copy-Item -LiteralPath (Join-Path $tmp ($name + '.pdf')) -Destination ([IO.Path]::ChangeExtension($outFull, '.pdf')) -Force }
  Write-Output "OK $Out"
}
catch { Write-Output ("RENDER ERROR: " + $_ + " @ " + $_.ScriptStackTrace); try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}; exit 1 }
finally {
  # Never close the user's own PowerPoint: quit only when no other presentation is open.
  $quit = $false
  try { if ($app.Presentations.Count -eq 0) { $app.Quit(); $quit = $true } } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  # Wait for PowerPoint to exit so the next job does not attach to an instance that is shutting down.
  if ($quit) { for ($w = 0; $w -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Milliseconds 500 } }
}
