param(
  [Parameter(Mandatory=$true)][string]$Spec,
  [Parameter(Mandatory=$true)][string]$Style,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$Brand,
  [string]$PngDir,
  [string]$HistoryPath,
  [switch]$Potx
)
# Builds a .pptx from a content spec (JSON), a layout style (JSON) and a brand (JSON) by driving PowerPoint over COM.
# Usage: powershell -ExecutionPolicy Bypass -File build.ps1 -Spec spec.json -Style style.json -Out deck.pptx [-Brand brand.json] [-PngDir dir] [-HistoryPath history.json] [-Potx]
$ErrorActionPreference = 'Stop'
$EngineDir = Split-Path $PSScriptRoot -Parent

function LongPath([string]$p) { $full = [IO.Path]::GetFullPath($p); if ($full.Length -gt 240 -and -not $full.StartsWith('\\?\')) { return '\\?\' + $full }; return $full }
function ReadJson([string]$p) { return ([IO.File]::ReadAllText((LongPath $p), [Text.Encoding]::UTF8) | ConvertFrom-Json) }
function Has($o, [string]$name) { return ($null -ne $o) -and ($o.PSObject.Properties.Name -contains $name) -and ($null -ne $o.$name) -and ("$($o.$name)" -ne '') }
function Get($o, [string]$name, $default) { if (Has $o $name) { return $o.$name }; return $default }

$deck = ReadJson $Spec
$sty  = ReadJson $Style
if (-not $Brand) { $Brand = Join-Path $EngineDir 'brand.json' }
if (Test-Path -LiteralPath (LongPath $Brand)) {
  $br = ReadJson $Brand
  foreach ($k in @('colors','headFont','bodyFont','logo')) { if (Has $br $k) { $sty | Add-Member -NotePropertyName $k -NotePropertyValue $br.$k -Force } }
}

function RGBc([string]$hex) {
  $hex = $hex.TrimStart('#')
  return [int]([Convert]::ToInt32($hex.Substring(0,2),16) + [Convert]::ToInt32($hex.Substring(2,2),16)*256 + [Convert]::ToInt32($hex.Substring(4,2),16)*65536)
}
# ObjectThemeColor values
$TX1=13; $BG1=14; $TX2=15; $BG2=16; $A1=5; $A2=6; $A3=7; $A4=8; $A5=9; $A6=10

# ---------- layout knobs ----------
$motif      = [string](Get $sty 'motif' 'rounded')
$MSHAPE     = @{ rounded = 5; square = 1; circle = 9 }[$motif]; if (-not $MSHAPE) { $MSHAPE = 5 }
$RADIUS     = [double](Get $sty 'radius' 0.12)
$coverKind  = [string](Get $sty 'cover' 'blocks')
$coverTitle = [string](Get $sty 'coverTitle' 'middle')
$mirror     = ([string](Get $sty 'mirror' 'false')) -eq 'True'
$titleAlign = [string](Get $sty 'titleAlign' 'left')
$kicker     = ([string](Get $sty 'kicker' 'false')) -eq 'True'
$contentBg  = [string](Get $sty 'contentBg' 'lt1')
$agendaKind = [string](Get $sty 'agenda' 'grid')
$cardsKind  = [string](Get $sty 'cards' 'outline')
$statsKind  = [string](Get $sty 'stats' 'plain')
$processKind= [string](Get $sty 'process' 'horizontal')
$titleSize  = [double](Get $sty 'titleSize' 28)
$lightCover = ($coverKind -eq 'minimal')
$CARD = $BG2; if ($contentBg -eq 'lt2') { $CARD = $BG1 }
$TA = 1; if ($titleAlign -eq 'center') { $TA = 2 }
$TITLE_Y = 34; $TITLE_H = 52; if ($kicker) { $TITLE_Y = 46; $TITLE_H = 46 }
$logo = $null; if (Has $sty 'logo') { $logo = $sty.logo }
# Images are copied into the build's temp folder first: PowerPoint runs as a separate process and
# reliably reads %TEMP%, while some folders (redirected or virtualized) may not be visible to it.
$script:staged = @{}
function Stage([string]$p) {
  if ($script:staged.ContainsKey($p)) { return $script:staged[$p] }
  $dst = Join-Path $tmp ('img-' + $script:staged.Count + [IO.Path]::GetExtension($p))
  Copy-Item -LiteralPath (LongPath $p) -Destination $dst -Force
  $script:staged[$p] = $dst
  return $dst
}
function LogoPath([string]$p) { if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $EngineDir $p }; return (Stage $p) }
function MX($x, $w) { if ($mirror) { return 960 - $x - $w } else { return $x } }

function GetPh($container, [int[]]$types) {
  foreach ($sh in $container.Shapes) { if ($sh.Type -eq 14) { if ($types -contains $sh.PlaceholderFormat.Type) { return $sh } } }
  return $null
}
function DelPh($container, [int[]]$types) {
  $del = @(); foreach ($sh in $container.Shapes) { if ($sh.Type -eq 14 -and ($types -contains $sh.PlaceholderFormat.Type)) { $del += $sh } }
  foreach ($d in $del) { $d.Delete() }
}
function Pos($sh, $l, $t, $w, $h) { $sh.Left=[single]$l; $sh.Top=[single]$t; $sh.Width=[single]$w; $sh.Height=[single]$h }

function AddText($slide, [string]$text, $l, $t, $w, $h, $size, $color, [bool]$bold=$false, [int]$align=1, [int]$anchor=1) {
  $s = $slide.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
  $tf = $s.TextFrame
  $tf.WordWrap = -1; $tf.AutoSize = 0
  $s.Top = [single]$t; $s.Height = [single]$h   # new text boxes shrink to one line; restore the intended box
  $tf.MarginLeft=0; $tf.MarginRight=0; $tf.MarginTop=0; $tf.MarginBottom=0
  $tf.VerticalAnchor = $anchor
  $tr = $tf.TextRange
  $tr.Text = $text
  $tr.Font.Size = [single]$size
  $tr.Font.Bold = [int]$bold * -1
  $tr.Font.Color.ObjectThemeColor = $color
  $tr.ParagraphFormat.Alignment = $align
  $tr.ParagraphFormat.SpaceWithin = 1.1
  $tr.ParagraphFormat.WordWrap = -1   # Korean: break lines between words, not inside them
  return $s
}
function AddBox($slide, $shapeType, $l, $t, $w, $h, $fill, $radius=$RADIUS) {
  $s = $slide.Shapes.AddShape($shapeType, [single]$l, [single]$t, [single]$w, [single]$h)
  if ($shapeType -eq 5) { $s.Adjustments.Item(1) = [single]$radius }
  $s.Fill.Solid()
  $s.Fill.ForeColor.ObjectThemeColor = $fill
  $s.Line.Visible = 0
  $s.Shadow.Visible = 0
  return $s
}
function Outline($slide, $shapeType, $l, $t, $w, $h, $color, $weight=1.5) {
  $s = $slide.Shapes.AddShape($shapeType, [single]$l, [single]$t, [single]$w, [single]$h)
  if ($shapeType -eq 5) { $s.Adjustments.Item(1) = [single]$RADIUS }
  $s.Fill.Visible = 0; $s.Line.Visible = -1; $s.Line.Weight = [single]$weight
  $s.Line.ForeColor.ObjectThemeColor = $color; $s.Shadow.Visible = 0
  return $s
}
function Badge($slide, [string]$num, $l, $t, $size, $fill, $txt) {
  $b = AddBox $slide $MSHAPE $l $t $size $size $fill ([math]::Min(0.3, $RADIUS * 1.8))
  $tf = $b.TextFrame
  $tf.MarginLeft=0; $tf.MarginRight=0; $tf.MarginTop=0; $tf.MarginBottom=0
  $tf.VerticalAnchor = 3
  $tr = $tf.TextRange
  $tr.Text = $num
  $tr.Font.Size = [single][int]($size * 0.36)
  $tr.Font.Bold = -1
  $tr.Font.Color.ObjectThemeColor = $txt
  $tr.ParagraphFormat.Alignment = 2
  return $b
}
function Line($slide, $x1, $y1, $x2, $y2, $weight=1) {
  # Thin rectangle instead of AddLine (AddLine throws InvalidCast over COM)
  if ($x1 -eq $x2) { $l = $x1 - $weight/2; $t = [math]::Min($y1,$y2); $w = $weight; $h = [math]::Abs($y2-$y1) }
  else { $l = [math]::Min($x1,$x2); $t = $y1 - $weight/2; $w = [math]::Abs($x2-$x1); $h = $weight }
  $ln = $slide.Shapes.AddShape(1, [single]$l, [single]$t, [single]$w, [single]$h)
  $ln.Line.Visible = 0; $ln.Shadow.Visible = 0
  $ln.Fill.Solid(); $ln.Fill.ForeColor.ObjectThemeColor = $BG2
  try { $ln.Fill.ForeColor.Brightness = [single]-0.15 } catch {}
  return $ln
}
# Insert a picture that fills the box (center crop).
function PlaceFill($slide, [string]$path, $l, $t, $w, $h) {
  $pic = $slide.Shapes.AddPicture((Stage $path), 0, -1, [single]$l, [single]$t)
  $pic.LockAspectRatio = 0
  $w0 = [double]$pic.Width; $h0 = [double]$pic.Height
  $r = [math]::Max($w / $w0, $h / $h0)
  $cx = ($w0 - $w / $r) / 2; $cy = ($h0 - $h / $r) / 2
  $pic.PictureFormat.CropLeft = [single]$cx; $pic.PictureFormat.CropRight = [single]$cx
  $pic.PictureFormat.CropTop = [single]$cy; $pic.PictureFormat.CropBottom = [single]$cy
  Pos $pic $l $t $w $h
  return $pic
}
function LogoDims([double]$hDefault, [double]$maxW, [string]$wKey, [string]$hKey) {
  $hh = [double](Get $logo $hKey $hDefault)
  $ww = [double](Get $logo $wKey 0)
  if ($ww -le 0) {
    $a = [double](Get $logo 'aspect' 5.4); if ($a -le 0) { $a = 5.4 }
    $ww = $hh * $a
    if ($ww -gt $maxW) { $ww = $maxW; $hh = $ww / $a }
  }
  return @($ww, $hh)
}
# Cover logo. On dark backgrounds use the white version, or a white plate behind the color logo.
function AddLogo($container, [bool]$onDark, $x, $y) {
  if (-not $logo) { return $false }
  $dims = LogoDims 26 200 'coverWidth' 'coverHeight'; $w = $dims[0]; $h = $dims[1]
  if ($onDark -and (Has $logo 'white')) { $null = $container.Shapes.AddPicture((LogoPath $logo.white), 0, -1, [single]$x, [single]$y, [single]$w, [single]$h); return $true }
  if (-not (Has $logo 'color')) { return $false }
  if ($onDark) { $null = AddBox $container 5 ($x - 10) ($y - 8) ($w + 20) ($h + 16) $BG1 0.3 }
  $null = $container.Shapes.AddPicture((LogoPath $logo.color), 0, -1, [single]$x, [single]$y, [single]$w, [single]$h)
  return $true
}
function Notes($slide, $d) { if (Has $d 'notes') { $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = [string]$d.notes } }
$script:curKicker = ''
function Footer($slide, $d) {
  $slide.HeadersFooters.Footer.Visible = -1
  $slide.HeadersFooters.Footer.Text = [string](Get $deck 'footer' (Get $deck 'title' ''))
  $slide.HeadersFooters.SlideNumber.Visible = -1
  if ($kicker) {
    $k = [string](Get $d 'kicker' $script:curKicker)
    if ($k) { $null = AddText $slide $k 48 26 864 16 11 $A1 $true $TA }
  }
}
function SetTitle($slide, [string]$text) { (GetPh $slide @(1,3)).TextFrame.TextRange.Text = $text }
function Num([int]$i) { return ('{0:D2}' -f $i) }

# Cover geometry
$COVER_TW = 540; if ($coverKind -eq 'split' -or $coverKind -eq 'minimal') { $COVER_TW = 520 }
$COVER_X = MX 58 $COVER_TW
$CT = @{ titleY = 170; subY = 314; subH = 60; metaY = 456 }
if ($coverTitle -eq 'low') { $CT = @{ titleY = 250; subY = 392; subH = 50; metaY = 470 } }
function CoverDeco($ct) {
  switch ($coverKind) {
    'circles' {
      $d1 = AddBox $ct 9 (MX 600 470) 170 470 470 $A2; $d1.Fill.Transparency = 0.25
      $null = AddBox $ct 9 (MX 770 150) 70 150 150 $A1
      $null = Outline $ct 9 (MX 660 100) 400 100 100 $A3
    }
    'split' {
      $null = AddBox $ct 1 (MX 620 340) 0 340 540 $A1
      $null = Outline $ct $MSHAPE (MX 680 220) 130 220 220 $BG1 2
      $null = AddBox $ct $MSHAPE (MX 820 100) 300 100 100 $TX2
    }
    'minimal' {
      $null = AddBox $ct 1 (MX 620 340) 0 340 540 $BG2
      $null = AddBox $ct $MSHAPE (MX 680 200) 150 200 200 $A1
      $d2 = AddBox $ct $MSHAPE (MX 800 110) 290 110 110 $A2; $d2.Fill.Transparency = 0.15
    }
    default {
      $null = AddBox $ct $MSHAPE (MX 650 210) 96 210 210 $A1
      $d2 = AddBox $ct $MSHAPE (MX 748 170) 236 170 170 $A2; $d2.Fill.Transparency = 0.15
      $null = Outline $ct $MSHAPE (MX 618 96) 330 96 96 $A3
    }
  }
}

$app = New-Object -ComObject PowerPoint.Application
$tmp = Join-Path $env:TEMP ('ppt-build-' + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Force $tmp | Out-Null
try {
  $pres = $app.Presentations.Add(-1)
  try { if ($app.Presentations.Count -le 1) { $app.WindowState = 2 } } catch {}
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $master = $pres.SlideMaster
  $master.Design.Name = [string](Get $sty 'name' 'Custom')

  # ---------- Theme ----------
  $c = $sty.colors
  $order = @('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
  $tcs = $master.Theme.ThemeColorScheme
  for ($i=0; $i -lt 12; $i++) { $tcs.Colors($i+1).RGB = (RGBc ([string]$c.($order[$i]))) }
  $fs = $master.Theme.ThemeFontScheme
  $fs.MajorFont.Item(1).Name = [string]$sty.headFont; $fs.MajorFont.Item(3).Name = [string]$sty.headFont
  $fs.MinorFont.Item(1).Name = [string]$sty.bodyFont; $fs.MinorFont.Item(3).Name = [string]$sty.bodyFont

  # ---------- Master ----------
  $master.Background.Fill.Solid()
  if ($contentBg -eq 'lt2') { $master.Background.Fill.ForeColor.ObjectThemeColor = $BG2 } else { $master.Background.Fill.ForeColor.ObjectThemeColor = $BG1 }
  DelPh $master @(16)
  $mt = GetPh $master @(1)
  Pos $mt 48 $TITLE_Y 864 $TITLE_H
  $mt.TextFrame.MarginLeft = 0; $mt.TextFrame.MarginRight = 0; $mt.TextFrame.VerticalAnchor = 3
  $r = $mt.TextFrame.TextRange
  $r.Font.Size = [single]$titleSize; $r.Font.Bold = -1; $r.Font.Color.ObjectThemeColor = $TX2
  $r.ParagraphFormat.Alignment = $TA
  $r.ParagraphFormat.WordWrap = -1
  $mb = GetPh $master @(2)
  Pos $mb 48 120 864 360
  $mb.TextFrame.MarginLeft = 0
  $pc = $mb.TextFrame.TextRange.Paragraphs().Count
  for ($i=1; $i -le $pc; $i++) {
    $p = $mb.TextFrame.TextRange.Paragraphs($i)
    $p.Font.Color.ObjectThemeColor = $TX1
    $p.Font.Size = [single][math]::Max(12, 18 - 2*$i)
    $p.ParagraphFormat.SpaceAfter = 8
    $p.ParagraphFormat.SpaceWithin = 1.15
    $p.ParagraphFormat.Bullet.Visible = -1
    if ($motif -eq 'circle') { $p.ParagraphFormat.Bullet.Character = 9679 } else { $p.ParagraphFormat.Bullet.Character = 9632 }
    $p.ParagraphFormat.Bullet.RelativeSize = 0.6
    $p.ParagraphFormat.Bullet.Font.Color.ObjectThemeColor = $A1
    $p.ParagraphFormat.WordWrap = -1
  }
  $footLeft = 62
  if ($logo -and (Has $logo 'color')) {
    $fd = LogoDims 13 110 'footerWidth' 'footerHeight'
    $null = $master.Shapes.AddPicture((LogoPath $logo.color), 0, -1, [single]48, [single](515 - $fd[1]/2), [single]$fd[0], [single]$fd[1])
    $footLeft = 48 + $fd[0] + 14
  } else {
    $null = AddBox $master $MSHAPE 48 508.5 7 7 $A1 0.2
  }
  $mf = GetPh $master @(15)
  Pos $mf $footLeft 505 (560 - $footLeft) 20
  $mf.TextFrame.MarginLeft = 0; $mf.TextFrame.VerticalAnchor = 3
  $mf.TextFrame.TextRange.Font.Size = 9; $mf.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A6
  $mf.TextFrame.TextRange.ParagraphFormat.Alignment = 1
  $mn = GetPh $master @(13)
  Pos $mn 852 505 60 20
  $mn.TextFrame.MarginRight = 0; $mn.TextFrame.VerticalAnchor = 3
  $mn.TextFrame.TextRange.Font.Size = 9; $mn.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A6
  $mn.TextFrame.TextRange.ParagraphFormat.Alignment = 3

  foreach ($lay in $master.CustomLayouts) { $t = GetPh $lay @(1); if ($t) { Pos $t 48 $TITLE_Y 864 $TITLE_H; $t.TextFrame.MarginLeft = 0 } }

  # Capture layouts before Duplicate() shifts indexes
  $L1 = $master.CustomLayouts.Item(1)
  $L2 = $master.CustomLayouts.Item(2)
  $L3 = $master.CustomLayouts.Item(3)
  $L6 = $master.CustomLayouts.Item(6)

  # ---------- Cover layout ----------
  $L1.DisplayMasterShapes = 0
  $L1.FollowMasterBackground = 0
  $L1.Background.Fill.Solid()
  if ($lightCover) { $L1.Background.Fill.ForeColor.ObjectThemeColor = $BG1 } else { $L1.Background.Fill.ForeColor.ObjectThemeColor = $TX2 }
  DelPh $L1 @(13,15,16)
  $t = GetPh $L1 @(3,1)
  Pos $t $COVER_X $CT.titleY $COVER_TW 130
  $t.TextFrame.MarginLeft = 0; $t.TextFrame.VerticalAnchor = 4
  $tr = $t.TextFrame.TextRange
  $tr.Font.Size = 40; $tr.Font.Bold = -1; $tr.ParagraphFormat.Alignment = 1; $tr.ParagraphFormat.SpaceWithin = 1.05; $tr.ParagraphFormat.WordWrap = -1
  if ($lightCover) { $tr.Font.Color.ObjectThemeColor = $TX2 } else { $tr.Font.Color.ObjectThemeColor = $BG1 }
  $st = GetPh $L1 @(4)
  Pos $st $COVER_X $CT.subY $COVER_TW $CT.subH
  $st.TextFrame.MarginLeft = 0; $st.TextFrame.VerticalAnchor = 1
  $st.TextFrame.TextRange.Font.Size = 18; $st.TextFrame.TextRange.ParagraphFormat.Alignment = 1
  if ($lightCover) { $st.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A6 } else { $st.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A3 }
  CoverDeco $L1
  if (-not (AddLogo $L1 (-not $lightCover) $COVER_X 56)) { $null = AddBox $L1 $MSHAPE $COVER_X ($CT.titleY - 30) 16 16 $A1 0.22 }
  $LC = $L1.Duplicate()

  # ---------- Section layout ----------
  $SX = MX 58 640; $SA = 1; if ($mirror) { $SA = 3 }
  $L3.DisplayMasterShapes = 0
  $L3.FollowMasterBackground = 0
  $L3.Background.Fill.Solid()
  if ($lightCover) { $L3.Background.Fill.ForeColor.ObjectThemeColor = $BG2 } else { $L3.Background.Fill.ForeColor.ObjectThemeColor = $TX2 }
  DelPh $L3 @(13,15,16)
  $t = GetPh $L3 @(1)
  Pos $t $SX 250 640 70
  $t.TextFrame.MarginLeft = 0; $t.TextFrame.MarginRight = 0; $t.TextFrame.VerticalAnchor = 4
  $t.TextFrame.TextRange.Font.Size = 36; $t.TextFrame.TextRange.Font.Bold = -1; $t.TextFrame.TextRange.ParagraphFormat.Alignment = $SA
  if ($lightCover) { $t.TextFrame.TextRange.Font.Color.ObjectThemeColor = $TX2 } else { $t.TextFrame.TextRange.Font.Color.ObjectThemeColor = $BG1 }
  $b = GetPh $L3 @(2,7)
  Pos $b $SX 332 640 50
  $b.TextFrame.MarginLeft = 0; $b.TextFrame.MarginRight = 0
  $b.TextFrame.TextRange.Font.Size = 16
  $b.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
  $b.TextFrame.TextRange.ParagraphFormat.Alignment = $SA
  if ($lightCover) { $b.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A6 } else { $b.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A3 }
  if ($coverKind -eq 'circles') {
    $d = AddBox $L3 9 (MX 760 300) 300 300 300 $A2; $d.Fill.Transparency = 0.25
    $null = AddBox $L3 9 (MX 840 80) 250 80 80 $A1
  } else {
    $null = AddBox $L3 $MSHAPE (MX 790 120) 330 120 120 $A1
    $null = AddBox $L3 $MSHAPE (MX 860 70) 400 70 70 $A2
  }
  $b = GetPh $L2 @(7,2); Pos $b 48 124 864 350

  # ================= SLIDES =================
  $script:n = 0
  $pendingCharts = @()
  $sectionMarks = @()
  $mutedOnDark = $A3; if ($lightCover) { $mutedOnDark = $A6 }

  foreach ($d in $deck.slides) {
    $type = [string]$d.type
    $script:n++
    $idx = $script:n
    switch ($type) {
      { $_ -eq 'cover' -or $_ -eq 'closing' } {
        $lay = $L1; if ($type -eq 'closing') { $lay = $LC }
        $s = $pres.Slides.AddSlide($idx, $lay)
        SetTitle $s ([string]$d.title)
        (GetPh $s @(4)).TextFrame.TextRange.Text = [string](Get $d 'subtitle' '')
        if (Has $d 'image') {
          # Photo cover: hide the layout decoration and fill the decoration side with the photo.
          $s.DisplayMasterShapes = 0
          $null = PlaceFill $s ([string]$d.image) (MX 600 360) 0 360 540
          $ax = 624; if ($mirror) { $ax = 268 }
          $null = AddBox $s $MSHAPE $ax 448 68 68 $A1
          $null = AddLogo $s (-not $lightCover) $COVER_X 56
        }
        if (Has $d 'meta') { $null = AddText $s ([string]$d.meta) $COVER_X $CT.metaY $COVER_TW 20 12 $mutedOnDark }
        $defName = 'Cover'; if ($type -eq 'closing') { $defName = 'Closing' }
        $sectionMarks += ,@($idx, [string](Get $d 'section' $defName))
      }
      'section' {
        $s = $pres.Slides.AddSlide($idx, $L3)
        SetTitle $s ([string]$d.title)
        (GetPh $s @(2,7)).TextFrame.TextRange.Text = [string](Get $d 'desc' '')
        $num = [string](Get $d 'number' '01')
        $null = AddText $s $num (MX 58 300) 150 300 96 72 $A1 $true $SA 4
        $script:curKicker = $num + '   ' + (([string]$d.title) -split "[\r\n]")[0]
        $sectionMarks += ,@($idx, [string](Get $d 'section' $d.title))
      }
      'bullets' {
        $s = $pres.Slides.AddSlide($idx, $L2); Footer $s $d
        SetTitle $s ([string]$d.title)
        $body = GetPh $s @(7,2)
        $hasVisual = (Has $d 'image') -or (Has $d 'imageText')
        $bx = 48; $vx = 508; if ($mirror) { $bx = 492; $vx = 48 }
        if ($hasVisual) { Pos $body $bx 124 420 340 }
        $lines = @($d.bullets)
        $body.TextFrame.TextRange.Text = (($lines | ForEach-Object { ([string]$_).TrimStart('>') }) -join "`r")
        for ($i=0; $i -lt $lines.Count; $i++) { if (([string]$lines[$i]).StartsWith('>')) { $body.TextFrame.TextRange.Paragraphs($i+1).IndentLevel = 2 } }
        if (Has $d 'image') {
          $null = PlaceFill $s ([string]$d.image) $vx 124 404 340
        } elseif (Has $d 'imageText') {
          $img = AddBox $s 5 $vx 124 404 340 $CARD 0.04
          $img.Line.Visible = -1; $img.Line.DashStyle = 4; $img.Line.Weight = 1; $img.Line.ForeColor.ObjectThemeColor = $A3
          $img.TextFrame.TextRange.Text = [string]$d.imageText
          $img.TextFrame.TextRange.Font.Size = 14
          $img.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A6
          $img.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
          $img.TextFrame.TextRange.Paragraphs(1).Font.Color.ObjectThemeColor = $TX2
        }
      }
      'photo' {
        $s = $pres.Slides.AddSlide($idx, $L6)
        $s.DisplayMasterShapes = 0
        $ix = MX 480 480; $tx = 48; if ($mirror) { $tx = 528 }
        if (Has $d 'image') { $null = PlaceFill $s ([string]$d.image) $ix 0 480 540 }
        else {
          $ph = AddBox $s 1 $ix 0 480 540 $BG2
          $ph.TextFrame.TextRange.Text = 'Photo'
          $ph.TextFrame.TextRange.Font.Size = 14; $ph.TextFrame.TextRange.Font.Color.ObjectThemeColor = $A6
        }
        $null = AddBox $s $MSHAPE $tx 112 14 14 $A1 0.22
        $t = GetPh $s @(1)
        Pos $t $tx 136 384 110
        $t.TextFrame.VerticalAnchor = 4
        $t.TextFrame.TextRange.Text = [string]$d.title
        $t.TextFrame.TextRange.ParagraphFormat.Alignment = 1
        if (Has $d 'text') { $null = AddText $s ([string]$d.text) $tx 262 384 170 16 $TX1 }
        if (Has $d 'caption') { $null = AddText $s ([string]$d.caption) $tx 470 384 20 11 $A6 }
      }
      'agenda' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string](Get $d 'title' 'Agenda'))
        $items = @($d.items); $cnt = $items.Count
        if ($agendaKind -eq 'list') {
          $rh = [math]::Min(68, 340 / [math]::Max(1,$cnt))
          for ($i=0; $i -lt $cnt; $i++) {
            $y = 124 + $i * $rh
            $null = AddText $s (Num ($i+1)) 48 $y 80 $rh 28 $A1 $true 1 3
            $null = AddText $s ([string]$items[$i].title) 136 $y 320 $rh 20 $TX2 $true 1 3
            if (Has $items[$i] 'desc') { $null = AddText $s ([string]$items[$i].desc) 470 $y 442 $rh 14 $A6 $false 1 3 }
            $null = Line $s 48 ($y + $rh) 912 ($y + $rh)
          }
        } else {
          $rows = [math]::Ceiling($cnt / 2)
          $h = [math]::Min(126, (340 - ($rows-1)*20) / [math]::Max(1,$rows))
          for ($i=0; $i -lt $cnt; $i++) {
            $x = 48 + ($i % 2) * 444; $y = 130 + [math]::Floor($i / 2) * ($h + 20)
            $null = AddBox $s 5 $x $y 420 $h $CARD 0.06
            $null = Badge $s (Num ($i+1)) ($x+28) ($y + ($h-48)/2) 48 $TX2 $BG1
            if (Has $items[$i] 'desc') {
              $null = AddText $s ([string]$items[$i].title) ($x+96) ($y + $h/2 - 26) 300 26 20 $TX2 $true
              $null = AddText $s ([string]$items[$i].desc) ($x+96) ($y + $h/2 + 6) 296 40 14 $A6
            } else {
              $null = AddText $s ([string]$items[$i].title) ($x+96) $y 300 $h 20 $TX2 $true 1 3
            }
          }
        }
      }
      'cards' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string]$d.title)
        if (Has $d 'lead') { $null = AddText $s ([string]$d.lead) 48 ($TITLE_Y + $TITLE_H + 4) 864 22 14 $A6 $false $TA }
        $cards = @($d.cards); $cnt = [math]::Max(1, $cards.Count)
        if ($cardsKind -eq 'stack') {
          $rh = [math]::Min(100, (330 - ($cnt-1)*14) / $cnt)
          for ($i=0; $i -lt $cards.Count; $i++) {
            $y = 136 + $i * ($rh + 14)
            $null = AddBox $s 5 48 $y 864 $rh $CARD 0.1
            $null = Badge $s (Num ($i+1)) 72 ($y + ($rh-44)/2) 44 $TX2 $BG1
            $null = AddText $s ([string]$cards[$i].title) 140 $y 280 $rh 18 $TX2 $true 1 3
            $null = AddText $s ([string]$cards[$i].desc) 440 $y 448 $rh 14 $A6 $false 1 3
          }
        } else {
          $w = (864 - ($cnt-1)*24) / $cnt
          for ($i=0; $i -lt $cards.Count; $i++) {
            $x = 48 + $i * ($w + 24)
            if ($cardsKind -eq 'filled') {
              $null = AddBox $s 5 $x 140 $w 280 $TX2 0.05
              $null = Badge $s (Num ($i+1)) ($x+28) 168 52 $A1 $BG1
              $null = AddText $s ([string]$cards[$i].title) ($x+28) 246 ($w-56) 28 20 $BG1 $true
              $null = AddText $s ([string]$cards[$i].desc) ($x+28) 286 ($w-56) 120 14 $A3
            } else {
              $cb = AddBox $s 5 $x 140 $w 280 $BG1 0.05
              if ($contentBg -eq 'lt2') { $cb.Line.Visible = 0 } else { $cb.Line.Visible = -1; $cb.Line.Weight = 1; $cb.Line.ForeColor.ObjectThemeColor = $BG2; try { $cb.Line.ForeColor.Brightness = [single]-0.1 } catch {} }
              $null = Badge $s (Num ($i+1)) ($x+28) 168 52 $TX2 $BG1
              $null = AddText $s ([string]$cards[$i].title) ($x+28) 246 ($w-56) 28 20 $TX2 $true
              $null = AddText $s ([string]$cards[$i].desc) ($x+28) 286 ($w-56) 120 14 $A6
            }
          }
        }
      }
      'stats' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string]$d.title)
        $stats = @($d.stats); $cnt = [math]::Max(1, $stats.Count)
        $w = (864 - ($cnt-1)*24) / $cnt
        $vs = 54; if ($cnt -ge 4) { $vs = 44 }
        for ($i=0; $i -lt $stats.Count; $i++) {
          $x = 48 + $i * ($w + 24)
          if ($statsKind -eq 'tiles') {
            $null = AddBox $s 5 $x 120 $w 180 $CARD 0.06
            $null = AddText $s ([string]$stats[$i].value) ($x+24) 136 ($w-48) 70 ($vs-6) $A1 $true 1 4
            $null = AddText $s ([string]$stats[$i].label) ($x+24) 216 ($w-48) 26 18 $TX2 $true
            if (Has $stats[$i] 'note') { $null = AddText $s ([string]$stats[$i].note) ($x+24) 246 ($w-48) 40 14 $A6 }
          } else {
            $null = AddText $s ([string]$stats[$i].value) $x 130 $w 80 $vs $TX2 $true 1 4
            $null = AddText $s ([string]$stats[$i].label) $x 222 $w 26 18 $TX1 $true
            if (Has $stats[$i] 'note') { $null = AddText $s ([string]$stats[$i].note) $x 252 $w 22 14 $A6 }
            if ($i -gt 0) { $null = Line $s ($x - 12) 140 ($x - 12) 270 }
          }
        }
        if (Has $d 'insight') {
          $null = AddBox $s 5 48 330 864 130 $TX2 0.06
          $null = AddText $s ([string](Get $d 'insightLabel' 'KEY INSIGHT')) 80 356 400 20 12 $A3 $true
          $null = AddText $s ([string]$d.insight) 80 384 800 56 18 $BG1 $true
        }
      }
      'chart' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string]$d.title)
        $cw = 864; $cx = 40
        if (Has $d 'insights') {
          $cw = 540; $ix = 612; if ($mirror) { $ix = 48; $cx = 372 }
          $ins = @($d.insights)
          for ($i=0; $i -lt $ins.Count; $i++) {
            $y = 140 + $i * (330 / [math]::Max(3, $ins.Count))
            $null = AddBox $s $MSHAPE $ix ($y+5) 12 12 $A1 0.22
            $null = AddText $s ([string]$ins[$i].title) ($ix+24) $y 276 24 17 $TX2 $true
            if (Has $ins[$i] 'desc') { $null = AddText $s ([string]$ins[$i].desc) ($ix+24) ($y+30) 276 50 14 $A6 }
          }
        }
        if (Has $d 'caption') { $null = AddText $s ([string]$d.caption) ($cx + 8) 112 $cw 20 13 $TX1 $true }
        $pendingCharts += ,@($idx, $d, $cw, $cx)
      }
      'process' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string]$d.title)
        $steps = @($d.steps); $cnt = [math]::Max(1, $steps.Count)
        $cur = [int](Get $d 'current' -1)
        if ($processKind -eq 'vertical') {
          $rh = [math]::Min(84, 340 / $cnt)
          if ($steps.Count -gt 1) { $null = Line $s 74 (124 + $rh/2) 74 (124 + ($steps.Count-1)*$rh + $rh/2) 2 }
          for ($i=0; $i -lt $steps.Count; $i++) {
            $y = 124 + $i * $rh
            $fill = $TX2; if ($i -eq $cur) { $fill = $A1 }
            $null = Badge $s (Num ($i+1)) 50 ($y + ($rh-48)/2) 48 $fill $BG1
            $null = AddText $s ('STEP ' + (Num ($i+1))) 124 ($y + $rh/2 - 24) 200 16 11 $A1 $true
            $null = AddText $s ([string]$steps[$i].title) 124 ($y + $rh/2 - 6) 300 28 20 $TX2 $true
            if (Has $steps[$i] 'desc') { $null = AddText $s ([string]$steps[$i].desc) 440 $y 330 $rh 14 $A6 $false 1 3 }
            if (Has $steps[$i] 'when') { $null = AddText $s ([string]$steps[$i].when) 780 $y 132 $rh 12 $A6 $true 3 3 }
            if ($i -lt $steps.Count - 1) { $null = Line $s 124 ($y + $rh) 912 ($y + $rh) }
          }
        } else {
          $colW = 864 / $cnt
          $null = Line $s 78 186 (48 + ($cnt-1)*$colW + 30) 186 2
          $anyWhen = $false
          for ($i=0; $i -lt $steps.Count; $i++) {
            $x = 48 + $i * $colW
            $fill = $TX2; if ($i -eq $cur) { $fill = $A1 }
            $null = Badge $s (Num ($i+1)) $x 156 60 $fill $BG1
            $null = AddText $s ('STEP ' + (Num ($i+1))) $x 240 ($colW-20) 18 11 $A1 $true
            $null = AddText $s ([string]$steps[$i].title) $x 262 ($colW-20) 30 22 $TX2 $true
            if (Has $steps[$i] 'desc') { $null = AddText $s ([string]$steps[$i].desc) $x 302 ($colW-20) 60 14 $A6 }
            if (Has $steps[$i] 'when') { $anyWhen = $true; $null = AddText $s ([string]$steps[$i].when) $x 386 ($colW-20) 20 12 $A6 $true }
          }
          if ($anyWhen) { $null = Line $s 48 370 912 370 }
        }
      }
      'compare' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string]$d.title)
        $null = AddBox $s 5 48 120 410 350 $CARD 0.05
        $null = AddBox $s 5 502 120 410 350 $TX2 0.05
        $null = AddText $s ([string](Get $d.left 'label' 'AS-IS')) 80 148 340 18 12 $A6 $true
        $null = AddText $s ([string]$d.left.title) 80 170 340 32 22 $TX2 $true
        $null = AddText $s ([string](Get $d.right 'label' 'TO-BE')) 534 148 340 18 12 $A3 $true
        $null = AddText $s ([string]$d.right.title) 534 170 340 32 22 $BG1 $true
        $li = @($d.left.items); $ri = @($d.right.items)
        $cnt = [math]::Max($li.Count, $ri.Count); $gap = [math]::Min(70, 210 / [math]::Max(1,$cnt))
        for ($i=0; $i -lt $cnt; $i++) {
          $y = 236 + $i * $gap
          if ($i -lt $li.Count) { $null = AddText $s ([string][char]0x2014) 80 $y 20 24 16 $A6 $true; $null = AddText $s ([string]$li[$i]) 106 $y 320 ($gap-6) 16 $TX1 }
          if ($i -lt $ri.Count) { $null = AddBox $s $MSHAPE 536 ($y+5) 12 12 $A1 0.22; $null = AddText $s ([string]$ri[$i]) 560 $y 320 ($gap-6) 16 $BG1 }
        }
        $ar = AddBox $s 9 462 277 36 36 $A1
        $ar.TextFrame.MarginLeft=0; $ar.TextFrame.MarginRight=0
        $ar.TextFrame.TextRange.Text = [string][char]0x2192
        $ar.TextFrame.TextRange.Font.Size = 16; $ar.TextFrame.TextRange.Font.Bold = -1
        $ar.TextFrame.TextRange.Font.Color.ObjectThemeColor = $BG1
      }
      'table' {
        $s = $pres.Slides.AddSlide($idx, $L6); Footer $s $d
        SetTitle $s ([string]$d.title)
        $hdr = @($d.headers); $rowsData = @($d.rows)
        $nr = $rowsData.Count + 1; $nc = $hdr.Count
        $rh = [math]::Min(56, 340 / $nr)
        $tshape = $s.Shapes.AddTable($nr, $nc, [single]48, [single]124, [single]864, [single]($rh * $nr))
        $tbl = $tshape.Table
        $tbl.ApplyStyle("{2D5ABB26-0587-4C30-8999-92F81FD0307C}")
        $first = [math]::Round(864 * 0.4); if ($nc -le 2) { $first = 432 }
        $rest = (864 - $first) / [math]::Max(1, $nc - 1)
        for ($cI=1; $cI -le $nc; $cI++) { if ($cI -eq 1) { $tbl.Columns.Item($cI).Width = [single]$first } else { $tbl.Columns.Item($cI).Width = [single]$rest } }
        $hl = @{}; if (Has $d 'statusColors') { foreach ($p in $d.statusColors.PSObject.Properties) { $hl[$p.Name] = $p.Value } }
        $tmap = @{ accent1=$A1; accent2=$A2; accent3=$A3; accent4=$A4; accent5=$A5; accent6=$A6; dk2=$TX2 }
        for ($rI=1; $rI -le $nr; $rI++) {
          $tbl.Rows.Item($rI).Height = [single]$rh
          for ($cI=1; $cI -le $nc; $cI++) {
            $cell = $tbl.Cell($rI, $cI)
            $tf = $cell.Shape.TextFrame
            $tf.VerticalAnchor = 3; $tf.MarginLeft = 16; $tf.MarginRight = 16
            $tr = $tf.TextRange
            if ($rI -eq 1) { $tr.Text = [string]$hdr[$cI-1] } else { $tr.Text = [string]@($rowsData[$rI-2])[$cI-1] }
            $tr.Font.Size = 14
            $cell.Shape.Fill.Visible = -1; $cell.Shape.Fill.Solid()
            if ($rI -eq 1) {
              $cell.Shape.Fill.ForeColor.ObjectThemeColor = $TX2
              $tr.Font.Bold = -1; $tr.Font.Color.ObjectThemeColor = $BG1
            } else {
              if ($rI % 2 -eq 0) { $cell.Shape.Fill.ForeColor.ObjectThemeColor = $BG1 } else { $cell.Shape.Fill.ForeColor.ObjectThemeColor = $BG2 }
              $tr.Font.Color.ObjectThemeColor = $TX1
              if ($cI -eq 1) { $tr.Font.Bold = -1 }
              if ($hl.ContainsKey($tr.Text) -and $tmap.ContainsKey([string]$hl[$tr.Text])) { $tr.Font.Bold = -1; $tr.Font.Color.ObjectThemeColor = $tmap[[string]$hl[$tr.Text]] }
            }
          }
        }
      }
      default { throw "Unknown slide type: $type" }
    }
    Notes $s $d
  }

  # ---------- Charts (last: the chart's Excel session disturbs later COM calls) ----------
  $ctype = @{ column = 51; bar = 57; line = 65; pie = 5 }
  foreach ($pc in $pendingCharts) {
    $s = $pres.Slides.Item($pc[0]); $d = $pc[1]; $cw = $pc[2]; $cx = $pc[3]
    $kind = [string](Get $d 'chartType' 'column')
    try {
      $chs = $s.Shapes.AddChart2(-1, $ctype[$kind], [single]$cx, [single]140, [single]($cw + 8), [single]330)
      $ch = $chs.Chart
      $ch.ChartData.Activate()
      $wb = $ch.ChartData.Workbook
      $ws = $wb.Worksheets.Item(1)
      $ws.Range("A1:D20").Clear()
      $cats = @($d.categories); $vals = @($d.values)
      $ws.Cells.Item(1,1).Value2 = ' '
      $ws.Cells.Item(1,2).Value2 = [string](Get $d 'seriesName' 'Value')
      for ($i=0; $i -lt $cats.Count; $i++) { $ws.Cells.Item($i+2,1).Value2 = [string]$cats[$i]; $ws.Cells.Item($i+2,2).Value2 = [double]$vals[$i] }
      $last = $cats.Count + 1
      try { $null = $ws.ListObjects.Item(1).Resize($ws.Range("A1:B$last")) } catch {}
      $ch.SetSourceData("='" + $ws.Name + "'!`$A`$1:`$B`$" + $last, 2)
      $wb.Close(); Start-Sleep -Seconds 3
      $ch.HasTitle = $true; $ch.ChartTitle.Delete(); $ch.HasTitle = $false
      $ch.ChartArea.Format.Line.Visible = 0
      $ch.ChartArea.Format.Fill.Visible = 0
      $ser = $ch.SeriesCollection(1)
      $ser.HasDataLabels = $true
      $dl = $ser.DataLabels()
      $dl.Format.TextFrame2.TextRange.Font.Size = 12
      $dl.Format.TextFrame2.TextRange.Font.Bold = -1
      if ($kind -eq 'pie') {
        $ch.HasLegend = $true; $ch.Legend.Position = -4107
        $ch.Legend.Format.TextFrame2.TextRange.Font.Size = 12
        $pal = @(3,5,6,7,8,10)
        for ($i=1; $i -le $cats.Count; $i++) { $ser.Points($i).Format.Fill.ForeColor.ObjectThemeColor = $pal[($i-1) % $pal.Count] }
        $dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = $BG1
      } else {
        $ch.HasLegend = $false
        $dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = $TX1
        if ($kind -eq 'line') {
          $ser.Format.Line.ForeColor.ObjectThemeColor = 3; $ser.Format.Line.Weight = 3
          $ser.MarkerStyle = 8; $ser.MarkerSize = 8
          $ser.MarkerBackgroundColor = (RGBc ([string]$c.accent1)); $ser.MarkerForegroundColor = (RGBc ([string]$c.accent1))
          $dl.Position = 0
        } else {
          $ch.ChartGroups(1).GapWidth = 70
          $ser.Format.Fill.Solid()
          $ser.Format.Fill.ForeColor.ObjectThemeColor = 3
          if ((Has $d 'highlight') -and [int]$d.highlight -ge 0 -and [int]$d.highlight -lt $cats.Count) { $ser.Points([int]$d.highlight + 1).Format.Fill.ForeColor.ObjectThemeColor = $A1 }
        }
        $va = $ch.Axes(2)
        $va.MajorGridlines.Format.Line.ForeColor.RGB = (RGBc 'E3E7EE')
        $va.Format.Line.Visible = 0
        $va.TickLabels.Font.Size = 10
        $va.TickLabels.Font.Color = (RGBc ([string]$c.accent6))
        $ca = $ch.Axes(1)
        $ca.TickLabels.Font.Size = 11
        $ca.TickLabels.Font.Color = (RGBc ([string]$c.dk1))
        $ca.Format.Line.ForeColor.RGB = (RGBc 'C9D0DB')
      }
    } catch { Write-Output ("CHART ERROR on slide $($pc[0]): " + $_) }
    Start-Sleep -Seconds 2
  }

  # ---------- Sections ----------
  $sp = $pres.SectionProperties
  $starts = @{}
  foreach ($m in $sectionMarks) { $starts[[int]$m[0]] = [string]$m[1] }
  if (-not $starts.ContainsKey(1)) { $starts[1] = [string](Get $deck 'firstSection' 'Intro') }
  if ($pres.Slides.Count -ge 2 -and -not $starts.ContainsKey(2) -and $starts[1] -eq 'Cover') { $starts[2] = [string](Get $deck 'overviewSection' 'Overview') }
  foreach ($k in ($starts.Keys | Sort-Object)) { $null = $sp.AddBeforeSlide([int]$k, $starts[$k]) }

  # ---------- Save ----------
  $name = [IO.Path]::GetFileNameWithoutExtension($Out)
  $tpptx = Join-Path $tmp ($name + '.pptx')
  $pres.SaveAs($tpptx, 24)
  if ($Potx) { $pres.SaveCopyAs((Join-Path $tmp ($name + '.potx')), 26) }
  if ($PngDir) {
    New-Item -ItemType Directory -Force (LongPath $PngDir) | Out-Null
    for ($i=1; $i -le $pres.Slides.Count; $i++) {
      $png = Join-Path $tmp ("slide-{0:D2}.png" -f $i)
      $pres.Slides.Item($i).Export($png, "PNG", 1280, 720)
      Copy-Item -LiteralPath $png -Destination (LongPath (Join-Path $PngDir ("slide-{0:D2}.png" -f $i))) -Force
    }
  }
  $pres.Close()
  $outFull = LongPath $Out
  New-Item -ItemType Directory -Force (Split-Path $outFull -Parent) | Out-Null
  Copy-Item -LiteralPath $tpptx -Destination $outFull -Force
  if ($Potx) { Copy-Item -LiteralPath (Join-Path $tmp ($name + '.potx')) -Destination ([IO.Path]::ChangeExtension($outFull, '.potx')) -Force }

  # ---------- History (keeps the next deck's layout different) ----------
  if (-not $HistoryPath) {
    $shared = Join-Path $env:APPDATA 'PPT-Editor'
    if (Test-Path $shared) { $HistoryPath = Join-Path $shared 'history.json' } else { $HistoryPath = Join-Path $EngineDir 'history.json' }
  }
  $hist = @()
  if (Test-Path $HistoryPath) { try { $hist = @([IO.File]::ReadAllText($HistoryPath, [Text.Encoding]::UTF8) | ConvertFrom-Json) } catch { $hist = @() } }
  $hist += [pscustomobject]@{
    date = (Get-Date -Format 'yyyy-MM-dd HH:mm'); deck = $name; style = [string](Get $sty 'name' '')
    cover = $coverKind; coverTitle = $coverTitle; mirror = $mirror; titleAlign = $titleAlign; kicker = $kicker; motif = $motif
    contentBg = $contentBg; agenda = $agendaKind; cards = $cardsKind; stats = $statsKind; process = $processKind
    dk2 = [string]$c.dk2; accent1 = [string]$c.accent1; headFont = [string]$sty.headFont
  }
  if ($hist.Count -gt 30) { $hist = $hist[($hist.Count-30)..($hist.Count-1)] }
  [IO.File]::WriteAllText($HistoryPath, (ConvertTo-Json -InputObject @($hist) -Depth 4), (New-Object Text.UTF8Encoding $false))
  Write-Output "OK $Out"
}
catch { Write-Output ("BUILD ERROR: " + $_ + " @ " + $_.ScriptStackTrace); exit 1 }
finally {
  # Never close the user's own PowerPoint: quit only when no other presentation is open.
  $quit = $false
  try { if ($app.Presentations.Count -eq 0) { $app.Quit(); $quit = $true } } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  # Wait for PowerPoint to exit so the next job does not attach to an instance that is shutting down.
  if ($quit) { for ($w = 0; $w -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Milliseconds 500 } }
}
