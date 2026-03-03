param(
  [string]$ZoneJsonPath = "data/zones.json",
  [string]$OutputRoot = "ZoneCards",
  [switch]$Overwrite,
  [switch]$SkipJsonWrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function Get-FolderNameForBiome {
  param([string]$Biome)

  switch ($Biome.ToLowerInvariant()) {
    "sewers" { return "Sewers" }
    "industrial" { return "Urban-Industrial" }
    "transit" { return "Transit" }
    "civic" { return "Civic" }
    "labs" { return "Labs" }
    "outskirts" { return "Outskirts" }
    default { return "Misc" }
  }
}

function Get-BiomeColors {
  param([string]$Biome)

  switch ($Biome.ToLowerInvariant()) {
    "sewers" { return @{ A = "#0C352E"; B = "#1A6B58"; Accent = "#9BE7C4" } }
    "industrial" { return @{ A = "#2D2F33"; B = "#54595F"; Accent = "#F6B463" } }
    "transit" { return @{ A = "#1C2735"; B = "#36506E"; Accent = "#9AD1FF" } }
    "civic" { return @{ A = "#2E2433"; B = "#5E4A66"; Accent = "#F3B6D6" } }
    "labs" { return @{ A = "#162B2A"; B = "#2A6561"; Accent = "#A9FFF7" } }
    "outskirts" { return @{ A = "#2E3220"; B = "#59653A"; Accent = "#E9E6A1" } }
    default { return @{ A = "#2D2D2D"; B = "#575757"; Accent = "#F0F0F0" } }
  }
}

function Convert-HexToColor {
  param([string]$Hex)

  $clean = $Hex.TrimStart('#')
  if ($clean.Length -ne 6) {
    throw "Invalid hex color: $Hex"
  }

  return [System.Drawing.Color]::FromArgb(
    255,
    [Convert]::ToInt32($clean.Substring(0, 2), 16),
    [Convert]::ToInt32($clean.Substring(2, 2), 16),
    [Convert]::ToInt32($clean.Substring(4, 2), 16)
  )
}

function Get-DeterministicShift {
  param([string]$Text)

  $sum = 0
  foreach ($ch in $Text.ToCharArray()) {
    $sum += [int][char]$ch
  }
  return ($sum % 100)
}

function New-ZoneCardImage {
  param(
    [string]$TileId,
    [string]$TileName,
    [string]$Biome,
    [string]$OutputPath
  )

  $width = 1417
  $height = 827

  $colors = Get-BiomeColors -Biome $Biome
  $baseA = Convert-HexToColor $colors.A
  $baseB = Convert-HexToColor $colors.B
  $accent = Convert-HexToColor $colors.Accent

  $shift = Get-DeterministicShift "$TileId|$Biome"
  $angle = [single](15 + ($shift % 45))

  $bmp = New-Object System.Drawing.Bitmap($width, $height)
  $gfx = [System.Drawing.Graphics]::FromImage($bmp)
  $gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $gfx.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

  try {
    $rect = [System.Drawing.Rectangle]::new(0, 0, [int]$width, [int]$height)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $baseA, $baseB, $angle)
    $gfx.FillRectangle($brush, $rect)
    $brush.Dispose()

    $stripeColor = [System.Drawing.Color]::FromArgb(28, $accent)
    $stripeBrush = New-Object System.Drawing.SolidBrush($stripeColor)
    for ($x = -$height; $x -lt ($width + $height); $x += 120) {
      $points = @(
        [System.Drawing.Point]::new([int]$x, 0),
        [System.Drawing.Point]::new([int]($x + 44), 0),
        [System.Drawing.Point]::new([int]($x + $height + 44), [int]$height),
        [System.Drawing.Point]::new([int]($x + $height), [int]$height)
      )
      $gfx.FillPolygon($stripeBrush, $points)
    }
    $stripeBrush.Dispose()

    $overlay = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(92, 10, 10, 10))
    $gfx.FillRectangle($overlay, 0, 0, $width, $height)
    $overlay.Dispose()

    $panelMargin = 42
    $panelHeight = 220
    $panelRect = [System.Drawing.Rectangle]::new(
      [int]$panelMargin,
      [int]($height - $panelHeight - $panelMargin),
      [int]($width - ($panelMargin * 2)),
      [int]$panelHeight
    )
    $panelBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(160, 0, 0, 0))
    $gfx.FillRectangle($panelBrush, $panelRect)
    $panelBrush.Dispose()

    $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(190, $accent), 6)
    $gfx.DrawRectangle($borderPen, 8, 8, $width - 16, $height - 16)
    $borderPen.Dispose()

    $titleFont = New-Object System.Drawing.Font("Arial", 56, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $metaFont = New-Object System.Drawing.Font("Arial", 28, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $titleBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(245, 255, 255, 255))
    $metaBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, $accent))

    $titleRect = [System.Drawing.RectangleF]::new(
      [single]($panelRect.X + 24),
      [single]($panelRect.Y + 20),
      [single]($panelRect.Width - 48),
      [single]120
    )
    $metaRect = [System.Drawing.RectangleF]::new(
      [single]($panelRect.X + 24),
      [single]($panelRect.Y + 130),
      [single]($panelRect.Width - 48),
      [single]70
    )

    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisWord
    $fmt.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap

    $gfx.DrawString($TileName, $titleFont, $titleBrush, $titleRect, $fmt)
    $gfx.DrawString(("{0}  |  {1}" -f $Biome.ToUpperInvariant(), $TileId), $metaFont, $metaBrush, $metaRect, $fmt)

    $fmt.Dispose()
    $titleFont.Dispose()
    $metaFont.Dispose()
    $titleBrush.Dispose()
    $metaBrush.Dispose()

    $dir = Split-Path -Parent $OutputPath
    if (-not (Test-Path $dir)) {
      New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
  }
  finally {
    $gfx.Dispose()
    $bmp.Dispose()
  }
}

if (-not (Test-Path $ZoneJsonPath)) {
  throw "Zone JSON not found: $ZoneJsonPath"
}

$raw = Get-Content $ZoneJsonPath -Raw | ConvertFrom-Json
if (-not $raw.tiles -or -not ($raw.tiles -is [System.Collections.IEnumerable])) {
  throw "Invalid JSON shape in $ZoneJsonPath. Expected { tiles: [] }."
}

$created = 0
$skipped = 0
$updatedPaths = 0

foreach ($tile in $raw.tiles) {
  $id = [string]$tile.id
  $name = if ([string]::IsNullOrWhiteSpace([string]$tile.name)) { $id } else { [string]$tile.name }
  $biome = if ([string]::IsNullOrWhiteSpace([string]$tile.biome)) { "unknown" } else { [string]$tile.biome }

  $folder = Get-FolderNameForBiome -Biome $biome

  if (-not $tile.PSObject.Properties.Name.Contains("image") -or [string]::IsNullOrWhiteSpace([string]$tile.image)) {
    $tile | Add-Member -NotePropertyName image -NotePropertyValue "ZoneCards/$folder/$id.png" -Force
    $updatedPaths++
  }

  $outputPath = [string]$tile.image
  if ([string]::IsNullOrWhiteSpace($outputPath)) {
    continue
  }

  $relativePath = $outputPath -replace '^ZoneCards/', ''
  $fsPath = Join-Path $OutputRoot ($relativePath -replace '/', [IO.Path]::DirectorySeparatorChar)

  if ((Test-Path $fsPath) -and -not $Overwrite) {
    $skipped++
    continue
  }

  New-ZoneCardImage -TileId $id -TileName $name -Biome $biome -OutputPath $fsPath
  $created++
}

if (-not $SkipJsonWrite) {
  $jsonOut = $raw | ConvertTo-Json -Depth 12
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText((Resolve-Path $ZoneJsonPath), ($jsonOut + [Environment]::NewLine), $utf8NoBom)
}

Write-Output "Updated image paths in JSON: $updatedPaths"
Write-Output "Images created: $created"
Write-Output "Images skipped (already existed): $skipped"
