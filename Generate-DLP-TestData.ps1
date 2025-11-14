<# 
.SYNOPSIS
  Generate synthetic files for DLP testing across three tiers:
    1) General (no PII)
    2) Light PII (<10 rows)
    3) Heavy PII (100+ rows)

.DESCRIPTION
  Produces CSV, JSON, XML, HTML, TXT for each tier. 
  If available: 
    - XLSX via ImportExcel module
    - DOCX via Word COM (Microsoft Word installed)
  All data is FAKE and randomly generated for testing purposes.

.PARAMETER OutputDir
  Root folder to create tier subfolders and files.

.PARAMETER HeavyCount
  Number of rows for the heavy PII set (default 150).

.PARAMETER LightCount
  Number of rows for the light PII set (default 8, must be <10 to meet your spec).

.EXAMPLE
  .\Generate-DLP-TestData.ps1 -OutputDir "C:\DLP-Lab" -HeavyCount 200
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$OutputDir,
  [int]$HeavyCount = 150,
  [int]$LightCount = 8
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---- Helpers ---------------------------------------------------------------

function New-FolderSafe {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function Get-RandomItem {
  param([object[]]$From)
  $From[(Get-Random -Minimum 0 -Maximum $From.Count)]
}

function New-FakeName {
  $first = @("Alex","Taylor","Jordan","Casey","Riley","Morgan","Jamie","Avery","Sam","Cameron","Kai","Eden","Shawn","Rowan","Jesse","Marin","Drew","Skye","Hayden","Kendall")
  $last  = @("Lee","Tan","Ng","Lim","Wong","Chan","Goh","Liu","Ho","Chua","Yap","Toh","Quek","Yeo","Pang","Ow","Koh","Foo","Cheng","Chew")
  "$(Get-RandomItem $first) $(Get-RandomItem $last)"
}

function New-FakeDomain { Get-RandomItem @("example.com","test.local","lab.internal","sample.net","dev.null") }

function New-FakeEmail {
  $name = (New-FakeName) -replace ' ', '.'
  ($name.ToLower() + "@" + (New-FakeDomain))
}

function New-FakePhone {
  # Generic Singapore-ish mobile pattern 8/9xxxxxxx (not real)
  $start = Get-RandomItem @("8","9")
  $rest = -join ((1..7) | ForEach-Object { Get-Random -Minimum 0 -Maximum 10 })
  "$start$rest"
}

function New-FakeAddress {
  $streets = @("Jalan Bukit","Orchard Rd","Serangoon Ave","Ang Mo Kio Ave","Bedok North Rd","Clementi Ave","Tampines St","Woodlands Ave","Jurong West St","Pasir Ris Dr")
  $blocks  = (Get-Random -Minimum 1 -Maximum 400).ToString()
  $unit    = (Get-Random -Minimum 1 -Maximum 80).ToString("00")
  "$blocks $(Get-RandomItem $streets), #$unit-$(Get-Random -Minimum 1 -Maximum 80), Singapore 0$(Get-Random -Minimum 10000 -Maximum 99999)"
}

function New-FakeNRIC {
  # Pattern-like SG NRIC/TIN (not real, checksum ignored): S1234567D
  $prefix = Get-RandomItem @("S","T","F","G")
  $digits = -join ((1..7) | ForEach-Object { Get-Random -Minimum 0 -Maximum 10 })
  $suffix = Get-RandomItem @("A","B","C","D","E","F","G","H","I","Z","J","K","L","M","N","P","Q","R","S","T","U","V","W","X","Y")
  "$prefix$digits$suffix"
}

function New-TestCardNumber {
  # Use well-known test card numbers (Luhn-valid, non-billable)
  Get-RandomItem @(
    "4111 1111 1111 1111", # Visa test
    "5555 5555 5555 4444", # MasterCard test
    "3782 822463 10005",   # Amex test
    "6011 1111 1111 1117", # Discover test
    "4000 0000 0000 0002"  # Another Visa test
  )
}

function New-PiiRow {
  [pscustomobject]@{
    Name        = New-FakeName
    Email       = New-FakeEmail
    Phone       = New-FakePhone
    NRIC        = New-FakeNRIC
    CreditCard  = New-TestCardNumber
    Address     = New-FakeAddress
    Notes       = "Customer onboarding form; contains multiple identifiers."
  }
}

function New-NonPiiRow {
  [pscustomobject]@{
    Title       = Get-RandomItem @("Quarterly Update","Release Notes","How-To Article","Team Newsletter","Project Summary","FAQ")
    Category    = Get-RandomItem @("General","Marketing","Engineering","Operations","HR","Support")
    Version     = ("v{0}.{1}.{2}" -f (Get-Random -Minimum 0 -Maximum 5), (Get-Random -Minimum 0 -Maximum 10), (Get-Random -Minimum 0 -Maximum 20))
    Summary     = "This is generic content without personal or sensitive identifiers."
    Tags        = "sample; demo; docs"
  }
}

function Write-AsCsv {
  param($Data, [string]$Path)
  $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Write-AsJson {
  param($Data, [string]$Path)
  ($Data | ConvertTo-Json -Depth 6) | Set-Content -Path $Path -Encoding UTF8
}

function Write-AsXml {
  param($Data, [string]$Path, [string]$RootName = "Records")
  $xml = $Data | ConvertTo-Xml -As String -Depth 6 -NoTypeInformation
  # Replace root with custom name (ConvertTo-Xml defaults to Objects/Object)
  $xml = $xml -replace "<Objects>", "<$RootName>" -replace "</Objects>", "</$RootName>"
  Set-Content -Path $Path -Value $xml -Encoding UTF8
}

function Write-AsHtml {
  param($Data, [string]$Path, [string]$Title)
  $html = $Data | ConvertTo-Html -Title $Title -PreContent "<h2>$Title</h2>" | Out-String
  Set-Content -Path $Path -Value $html -Encoding UTF8
}

function Write-AsTxt {
  param($Data, [string]$Path, [string]$Header)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine($Header)
  [void]$sb.AppendLine(('-' * [Math]::Min(120, $Header.Length + 10)))
  foreach ($row in $Data) {
    foreach ($p in $row.PSObject.Properties) {
      [void]$sb.AppendLine(("{0}: {1}" -f $p.Name, $p.Value))
    }
    [void]$sb.AppendLine("")
  }
  Set-Content -Path $Path -Value $sb.ToString() -Encoding UTF8
}

function Write-AsXlsx {
  param($Data, [string]$Path, [string]$WorksheetName)
  # Requires ImportExcel module (https://github.com/dfinke/ImportExcel)
  if (-not (Get-Module -ListAvailable -Name ImportExcel)) { return $false }
  try {
    $Data | Export-Excel -Path $Path -WorksheetName $WorksheetName -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow
    return $true
  } catch {
    return $false
  }
}

function Write-AsDocx {
  param($Data, [string]$Path, [string]$Title)
  try {
    $word = New-Object -ComObject Word.Application
  } catch {
    return $false
  }
  try {
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    $selection.Style = "Heading 1"
    $selection.TypeText($Title)
    $selection.TypeParagraph()
    $selection.Style = "Normal"

    foreach ($row in $Data) {
      foreach ($p in $row.PSObject.Properties) {
        $selection.TypeText("{0}: {1}" -f $p.Name, $p.Value)
        $selection.TypeParagraph()
      }
      $selection.TypeParagraph()
    }

    $doc.SaveAs([ref]$Path)
    $doc.Close()
    $word.Quit()
    return $true
  } catch {
    try { if ($doc) { $doc.Close() } if ($word) { $word.Quit() } } catch {}
    return $false
  }
}

# ---- Build data sets -------------------------------------------------------

if ($LightCount -ge 10) {
  throw "LightCount must be < 10 to satisfy the 'less than 10' requirement."
}
if ($HeavyCount -lt 100) {
  throw "HeavyCount must be >= 100 to satisfy the 'more than 100' requirement."
}

$nowStamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$root = (Resolve-Path -LiteralPath $OutputDir).Path
New-FolderSafe $root

$tiers = @{
  "01-General_NoPII" = 1
  "02-Light_PII"     = 2
  "03-Heavy_PII"     = 3
}

foreach ($tierName in $tiers.Keys) {
  $tierPath = Join-Path $root $tierName
  New-FolderSafe $tierPath

  switch ($tiers[$tierName]) {
    1 {
      # General content, no PII
      $rows = 30
      $data = for ($i=1; $i -le $rows; $i++) { New-NonPiiRow }
    }
    2 {
      # Light PII (<10 rows)
      $data = for ($i=1; $i -le $LightCount; $i++) { New-PiiRow }
    }
    3 {
      # Heavy PII (100+ rows)
      $data = for ($i=1; $i -le $HeavyCount; $i++) { New-PiiRow }
    }
  }

  $base = "$($tierName)_$nowStamp"

  # CSV
  Write-AsCsv -Data $data -Path (Join-Path $tierPath "$base.csv")

  # JSON
  Write-AsJson -Data $data -Path (Join-Path $tierPath "$base.json")

  # XML
  Write-AsXml  -Data $data -Path (Join-Path $tierPath "$base.xml") -RootName "Records"

  # HTML
  Write-AsHtml -Data $data -Path (Join-Path $tierPath "$base.html") -Title $tierName

  # TXT
  Write-AsTxt  -Data $data -Path (Join-Path $tierPath "$base.txt") -Header "$tierName (Generated $(Get-Date))"

  # XLSX (optional)
  $xlsxPath = (Join-Path $tierPath "$base.xlsx")
  $xlsxOK = Write-AsXlsx -Data $data -Path $xlsxPath -WorksheetName ($tierName -replace '^\d{2}-','')
  if (-not $xlsxOK) {
    # leave a hint file if XLSX couldn't be created
    Set-Content -Path (Join-Path $tierPath "README_XLSX.txt") -Encoding UTF8 -Value @"
To generate .xlsx files, install ImportExcel module:
  Install-Module ImportExcel -Scope CurrentUser
Then rerun the script.
"@
  }

  # DOCX (optional)
  $docxPath = (Join-Path $tierPath "$base.docx")
  $docxOK = Write-AsDocx -Data $data -Path $docxPath -Title $tierName
  if (-not $docxOK) {
    Set-Content -Path (Join-Path $tierPath "README_DOCX.txt") -Encoding UTF8 -Value @"
To generate .docx files, run on Windows with Microsoft Word installed.
"@
  }
}

Write-Host "`nDone. Test data created under: $root" -ForegroundColor Green
