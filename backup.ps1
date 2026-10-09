# <#
# .SYNOPSIS
#     Downloads SDS PDF files from Excel links to a temporary folder, parses each PDF 
#     to extract the actual Revision Date directly from the SDS document content, 
#     renames the file as "DocumentID_mm-dd-yyyy.pdf", and moves it to the Renewals folder.
# #>

# param(
#     [int]$Limit = 0 # 0 means process all files
# )

# $ProgressPreference = 'SilentlyContinue'

# # Paths
# $excelFile = "c:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF\Book1.xlsx"
# $outputFolder = "c:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF\Renewals"
# $tempDownloadDir = "$env:TEMP\sds_temp_downloads"
# $tempCopy = "$env:TEMP\temp_Book1.xlsx"
# $extractPath = "$env:TEMP\excel_extracted"

# # Ensure directories exist
# if (-not (Test-Path $outputFolder)) { New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null }
# if (-not (Test-Path $tempDownloadDir)) { New-Item -ItemType Directory -Path $tempDownloadDir -Force | Out-Null }

# # Load iTextSharp & BouncyCastle DLLs for PDF text extraction
# Add-Type -AssemblyName System.IO.Compression.FileSystem

# $bcDir = "$env:TEMP\BouncyCastle"
# $bcZip = "$env:TEMP\BouncyCastle.zip"
# if (-not (Test-Path "$bcDir\lib\BouncyCastle.Crypto.dll")) {
#     $wc = New-Object System.Net.WebClient
#     $wc.DownloadFile("https://www.nuget.org/api/v2/package/BouncyCastle/1.8.9", $bcZip)
#     Expand-Archive -Path $bcZip -DestinationPath $bcDir -Force
# }
# $bcDll = Get-ChildItem -Path $bcDir -Filter "BouncyCastle.Crypto.dll" -Recurse | Select-Object -First 1 -ExpandProperty FullName
# Add-Type -Path $bcDll

# $iTextDir = "$env:TEMP\iTextSharp"
# $iTextZip = "$env:TEMP\iTextSharp.zip"
# if (-not (Test-Path "$iTextDir\lib\itextsharp.dll")) {
#     $wc = New-Object System.Net.WebClient
#     $wc.DownloadFile("https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3", $iTextZip)
#     Expand-Archive -Path $iTextZip -DestinationPath $iTextDir -Force
# }
# $iTextDll = "$iTextDir\lib\itextsharp.dll"
# Add-Type -Path $iTextDll

# # Function to extract full text from a PDF file
# function Extract-PdfText($pdfPath) {
#     try {
#         $reader = New-Object iTextSharp.text.pdf.PdfReader($pdfPath)
#         $text = ""
#         for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
#             $strategy = New-Object iTextSharp.text.pdf.parser.LocationTextExtractionStrategy
#             $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i, $strategy)
#             $text += "=== PAGE $i ===`n" + $pageText + "`n"
#         }
#         $reader.Close()
#         return $text
#     } catch {
#         return ""
#     }
# }

# # Month mapping for English & French SDS documents
# $monthMap = @{
#     "jan"="01"; "january"="01"; "janvier"="01"
#     "feb"="02"; "february"="02"; "fevrier"="02"; "fvrier"="02"
#     "mar"="03"; "march"="03"; "mars"="03"
#     "apr"="04"; "april"="04"; "avril"="04"
#     "may"="05"; "mai"="05"
#     "jun"="06"; "june"="06"; "juin"="06"
#     "jul"="07"; "july"="07"; "juillet"="07"
#     "aug"="08"; "august"="08"; "aout"="08"; "aot"="08"; "août"="08"
#     "sep"="09"; "sept"="09"; "september"="09"; "septembre"="09"
#     "oct"="10"; "october"="10"; "octobre"="10"
#     "nov"="11"; "november"="11"; "novembre"="11"
#     "dec"="12"; "december"="12"; "decembre"="12"; "dcembre"="12"
# }

# # Decision logic to find and parse Revision Date from SDS text
# function Parse-SdsRevisionDate($pdfText) {
#     if (-not $pdfText) { return $null }

#     # Patterns matching Revision Date / Effective Date / Issue Date
#     $revPatterns = @(
#         "(?:Revision\s*Date|Date\s*de\s*r[e]vision|Revised|Revision)\s*[:\-]?\s*([A-Za-z0-9\s,\/\-]+)",
#         "(?:Date\s*of\s*Issue|Issue\s*Date|Effective\s*Date|Date\s*of\s*Preparation)\s*[:\-]?\s*([A-Za-z0-9\s,\/\-]+)"
#     )

#     foreach ($pattern in $revPatterns) {
#         $matches = [regex]::Matches($pdfText, $pattern, "IgnoreCase")
#         foreach ($m in $matches) {
#             $snippet = $m.Groups[1].Value.Trim()
            
#             # Text dates like "29 août 2022", "August 29, 2022", "29-Aug-2022"
#             $textDateMatch = [regex]::Match($snippet, "(\d{1,2})?\s*([A-Za-z]+)\s*,?\s*(\d{1,2})?,?\s*(\d{4})")
#             if ($textDateMatch.Success) {
#                 $day1 = $textDateMatch.Groups[1].Value
#                 $monthStr = $textDateMatch.Groups[2].Value.ToLower()
#                 $day2 = $textDateMatch.Groups[3].Value
#                 $year = $textDateMatch.Groups[4].Value
                
#                 $monthNum = ""
#                 foreach ($k in $monthMap.Keys) {
#                     if ($monthStr -like "$k*") {
#                         $monthNum = $monthMap[$k]
#                         break
#                     }
#                 }
                
#                 if ($monthNum -and $year) {
#                     $day = if ($day1) { $day1 } else { $day2 }
#                     if (-not $day) { $day = "01" }
#                     $day = $day.PadLeft(2, '0')
#                     return "$monthNum-$day-$year"
#                 }
#             }
            
#             # Numeric dates MM/DD/YYYY or YYYY-MM-DD or MM-DD-YYYY
#             $numDateMatch = [regex]::Match($snippet, "(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})")
#             if ($numDateMatch.Success) {
#                 $mNum = $numDateMatch.Groups[1].Value.PadLeft(2, '0')
#                 $dNum = $numDateMatch.Groups[2].Value.PadLeft(2, '0')
#                 $yNum = $numDateMatch.Groups[3].Value
#                 if ($yNum.Length -eq 2) { $yNum = "20" + $yNum }
#                 return "$mNum-$dNum-$yNum"
#             }
#         }
#     }
#     return $null
# }

# # Parse Excel Sheet for URLs and Document IDs
# Copy-Item -Path $excelFile -Destination $tempCopy -Force
# if (Test-Path $extractPath) { Remove-Item -Recurse -Force $extractPath }
# [System.IO.Compression.ZipFile]::ExtractToDirectory($tempCopy, $extractPath)

# $sharedStrings = @()
# if (Test-Path "$extractPath\xl\sharedStrings.xml") {
#     [xml]$ssXml = Get-Content "$extractPath\xl\sharedStrings.xml"
#     foreach ($si in $ssXml.sst.si) {
#         if ($si.t) {
#             $sharedStrings += $si.t
#         } elseif ($si.r) {
#             $str = ($si.r | ForEach-Object { $_.t }) -join ""
#             $sharedStrings += $str
#         } else {
#             $sharedStrings += ""
#         }
#     }
# }

# $hyperlinks = @{}
# if (Test-Path "$extractPath\xl\worksheets\_rels\sheet1.xml.rels") {
#     [xml]$relsXml = Get-Content "$extractPath\xl\worksheets\_rels\sheet1.xml.rels"
#     foreach ($rel in $relsXml.Relationships.Relationship) {
#         if ($rel.TargetMode -eq "External") {
#             $hyperlinks[$rel.Id] = $rel.Target
#         }
#     }
# }

# [xml]$sheetXml = Get-Content "$extractPath\xl\worksheets\sheet1.xml"

# $cellHyperlinks = @{}
# if ($sheetXml.worksheet.hyperlinks) {
#     foreach ($hl in $sheetXml.worksheet.hyperlinks.hyperlink) {
#         $cellRef = $hl.ref
#         $rId = $hl.getAttribute("r:id")
#         if (-not $rId) { $rId = $hl.'id' }
#         if ($hyperlinks.ContainsKey($rId)) {
#             $cellHyperlinks[$cellRef] = $hyperlinks[$rId]
#         }
#     }
# }

# $items = @()
# foreach ($row in $sheetXml.worksheet.sheetData.row) {
#     $rowNum = [int]$row.r
#     if ($rowNum -eq 1) { continue } # Skip header
    
#     $cells = @{}
#     foreach ($c in $row.c) {
#         $colLetter = $c.r -replace '[0-9]', ''
#         $val = $c.v
#         if ($c.t -eq "s" -and $val -ne $null) {
#             $val = $sharedStrings[[int]$val]
#         }
#         $cells[$colLetter] = $val
#     }
    
#     $docId = $cells["A"]
#     $excelDateRaw = $cells["G"] # Fallback date
#     $url = $cellHyperlinks["M$rowNum"]
#     if (-not $url) { $url = $cells["M"] }
#     if (-not $url) { $url = $cellHyperlinks["W$rowNum"] }
#     if (-not $url) { $url = $cells["W"] }
    
#     if (-not $docId -or -not $url) { continue }
    
#     $fallbackDate = ""
#     if ($excelDateRaw -match '^\d+(\.\d+)?$') {
#         $fallbackDate = [DateTime]::FromOADate([double]$excelDateRaw).ToString("MM-dd-yyyy")
#     } else {
#         $fallbackDate = $excelDateRaw
#     }
    
#     $items += [PSCustomObject]@{
#         Row = $rowNum
#         DocID = $docId
#         FallbackDate = $fallbackDate
#         URL = $url
#     }
# }

# Write-Host "Found $($items.Count) records to process."

# if ($Limit -gt 0) {
#     $items = $items | Select-Object -First $Limit
# }

# $webClient = New-Object System.Net.WebClient
# $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

# $successCount = 0
# $failCount = 0

# foreach ($item in $items) {
#     $tempFile = Join-Path $tempDownloadDir "$($item.DocID)_temp.pdf"
    
#     Write-Host "Row $($item.Row) [DocID: $($item.DocID)] Downloading from $($item.URL) ..." -NoNewline
#     try {
#         # 1. Download to temp folder
#         $webClient.DownloadFile($item.URL, $tempFile)
        
#         # 2. Extract text from downloaded PDF
#         $pdfText = Extract-PdfText $tempFile
        
#         # 3. Parse revision date from PDF text
#         $sdsDate = Parse-SdsRevisionDate $pdfText
        
#         # 4. Fallback if PDF text didn't yield date
#         $finalDate = if ($sdsDate) { $sdsDate } else { $item.FallbackDate }
#         $dateSource = if ($sdsDate) { "SDS PDF" } else { "Excel Fallback" }
        
#         # 5. Save with nomenclature Documentid_mm-dd-yyyy.pdf
#         $finalFileName = "$($item.DocID)_${finalDate}.pdf"
#         $finalDestPath = Join-Path $outputFolder $finalFileName
        
#         Move-Item -Path $tempFile -Destination $finalDestPath -Force
        
#         Write-Host " [SUCCESS] Date: $finalDate (Source: $dateSource) -> Saved as $finalFileName" -ForegroundColor Green
#         $successCount++
#     } catch {
#         Write-Host " [FAILED]: $_" -ForegroundColor Red
#         if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
#         $failCount++
#     }
# }

# Write-Host "`nCompleted! Successfully processed: $successCount, Failed: $failCount"


























































<#
.SYNOPSIS
    Downloads SDS PDF files from Excel links, extracts the Revision Date directly
    from the PDF content (multi-language: English, French, Spanish, German, Italian,
    Portuguese, Dutch).
    - Files WITH a revision date  -> saved to Renewals as "DocumentID_mm-dd-yyyy.pdf"
    - Files WITHOUT a revision date -> saved to Renewals\no_revisions as "DocumentID.pdf"
    - Generates: download_report.xlsx and no_revisions\no_revisions_report.xlsx
#>

param(
    [int]$Limit = 0  # 0 means process all files
)

$ProgressPreference = 'SilentlyContinue'

# =============================================================================
# PATHS
# =============================================================================
$baseDir      = "c:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF"
$excelFile    = "$baseDir\Book1.xlsx"
$outputFolder = "$baseDir\Renewals"
$noRevFolder  = "$baseDir\Renewals\no_revisions"
$tempDownload = "$env:TEMP\sds_temp_downloads"
$tempCopy     = "$env:TEMP\temp_Book1.xlsx"
$extractPath  = "$env:TEMP\excel_extracted"

foreach ($dir in @($outputFolder, $noRevFolder, $tempDownload)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

# =============================================================================
# LOAD PDF LIBRARIES (iTextSharp + BouncyCastle)
# =============================================================================
Add-Type -AssemblyName System.IO.Compression.FileSystem

$bcDir = "$env:TEMP\BouncyCastle"
if (-not (Test-Path "$bcDir\lib\BouncyCastle.Crypto.dll")) {
    Write-Host "Downloading BouncyCastle..." -ForegroundColor Cyan
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile("https://www.nuget.org/api/v2/package/BouncyCastle/1.8.9", "$env:TEMP\BouncyCastle.zip")
    Expand-Archive -Path "$env:TEMP\BouncyCastle.zip" -DestinationPath $bcDir -Force
}
$bcDll = Get-ChildItem -Path $bcDir -Filter "BouncyCastle.Crypto.dll" -Recurse |
    Select-Object -First 1 -ExpandProperty FullName
Add-Type -Path $bcDll

$iTextDir = "$env:TEMP\iTextSharp"
if (-not (Test-Path "$iTextDir\lib\itextsharp.dll")) {
    Write-Host "Downloading iTextSharp..." -ForegroundColor Cyan
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile("https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3", "$env:TEMP\iTextSharp.zip")
    Expand-Archive -Path "$env:TEMP\iTextSharp.zip" -DestinationPath $iTextDir -Force
}
Add-Type -Path "$iTextDir\lib\itextsharp.dll"

# =============================================================================
# FUNCTION: Extract full text from a PDF file using iTextSharp
# =============================================================================
function Extract-PdfText([string]$pdfPath) {
    try {
        $reader = New-Object iTextSharp.text.pdf.PdfReader($pdfPath)
        $text = ""
        for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
            $strategy = New-Object iTextSharp.text.pdf.parser.LocationTextExtractionStrategy
            $text += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i, $strategy) + "`n"
        }
        $reader.Close()
        return $text
    }
    catch { return "" }
}

# =============================================================================
# FUNCTION: Resolve month name (multi-language) to 2-digit number
# Handles: English, French, Spanish, German, Italian, Portuguese, Dutch
# Uses prefix matching to handle accent-stripped variants
# =============================================================================
function Get-MonthNumber([string]$monthStr) {
    # Strip accents/diacritics and non-alpha chars, lowercase
    $s = $monthStr.ToLower() -replace '[^a-z]', ''

    switch -Wildcard ($s) {
        # ---- January group ----
        "jan*"    { return "01" }   # jan, january, janvier, januar, gennaio, janeiro, januari
        "ene*"    { return "01" }   # enero (Spanish)
        # ---- February group ----
        "feb*"    { return "02" }   # february, februar, febbraio, fevereiro, februari
        "fev*"    { return "02" }   # fevrier / fvrier (French)
        # ---- March group ----
        "mar*"    { return "03" }   # march, mars, marzo, maart, marco, marz, marz
        # ---- April group ----
        "apr*"    { return "04" }   # april, aprile
        "avr*"    { return "04" }   # avril (French)
        "abr*"    { return "04" }   # abril (Spanish/Portuguese)
        # ---- May group ----
        "may"     { return "05" }   # may (English)
        "mai"     { return "05" }   # mai (French/German)
        "mayo"    { return "05" }   # mayo (Spanish)
        "mag*"    { return "05" }   # maggio (Italian)
        "mei"     { return "05" }   # mei (Dutch)
        # ---- June group ----
        "jun*"    { return "06" }   # june, juin, junio, junho, juni, giugno
        "giu*"    { return "06" }   # giugno (Italian)
        # ---- July group ----
        "jul*"    { return "07" }   # july, juillet, julio, julho, juli
        "lug*"    { return "07" }   # luglio (Italian)
        # ---- August group ----
        "aug*"    { return "08" }   # august, augustus
        "ago*"    { return "08" }   # agosto (Spanish/Italian/Portuguese)
        "aou*"    { return "08" }   # aout / aout (French)
        "aot"     { return "08" }   # aot (corrupted French)
        # ---- September group ----
        "sep*"    { return "09" }   # september, septembre, septiembre, setembro
        "set*"    { return "09" }   # setiembre (Spanish), settembre (Italian)
        # ---- October group ----
        "oct*"    { return "10" }   # october, octobre
        "okt*"    { return "10" }   # oktober (German/Dutch)
        "out*"    { return "10" }   # outubro (Portuguese)
        "ott*"    { return "10" }   # ottobre (Italian)
        "oct*"    { return "10" }   # octubre (Spanish)
        # ---- November group ----
        "nov*"    { return "11" }   # november, novembre, noviembre, novembro
        # ---- December group ----
        "dec*"    { return "12" }   # december, decembre
        "dez*"    { return "12" }   # dezember (German), dezembro (Portuguese)
        "dic*"    { return "12" }   # diciembre (Spanish), dicembre (Italian)
        "dce*"    { return "12" }   # dcembre (corrupted French)
        default   { return $null }
    }
}

# =============================================================================
# MULTI-LANGUAGE REVISION DATE KEYWORD PATTERNS
# =============================================================================
$revisionKeywords = @(
    # English
    "Revision\s*Date",     "Date\s*of\s*Revision",
    "Revised",             "Issue\s*Date",
    "Date\s*of\s*Issue",   "Effective\s*Date",
    "Date\s*of\s*Preparation", "Prepared\s*Date",
    "Date\s*Prepared",     "Print\s*Date",

    # French
    "Date\s*de\s*r[eE]vision", "Date\s*d.?[eE]mission",
    "Date\s*de\s*pr[eE]paration", "Date\s*de\s*r[eE]daction",

    # Spanish
    "Fecha\s*de\s*Revisi.n",  "Fecha\s*de\s*Emisi.n",
    "Fecha\s*de\s*Preparaci.n", "Fecha\s*de\s*Publicaci.n",
    "Fecha\s*de\s*elaboraci.n",

    # German
    "Ausgabedatum",  "Revisionsdatum",  "Erstellungsdatum",
    "Datum\s*der\s*berarbeitung",       "Druckdatum",

    # Italian
    "Data\s*di\s*revisione",  "Data\s*di\s*emissione",
    "Data\s*di\s*preparazione",

    # Portuguese
    "Data\s*de\s*revis.o",    "Data\s*de\s*emiss.o",
    "Data\s*de\s*prepara..o",

    # Dutch
    "Revisiedatum",  "Herzieningsdatum",
    "Uitgiftedatum", "Datum\s*van\s*herziening"
)
$combinedKeyword = "(?:" + ($revisionKeywords -join "|") + ")"

# =============================================================================
# FUNCTION: Parse Revision Date from SDS text (multi-language)
# =============================================================================
function Parse-SdsRevisionDate([string]$pdfText) {
    if (-not $pdfText) { return $null }

    $patterns = @(
        "$combinedKeyword\s*[:\-]?\s*([A-Za-z0-9\s,\.\/\-]+)",
        "$combinedKeyword[\s\r\n]+([A-Za-z0-9\s,\.\/\-]+)"
    )

    foreach ($pat in $patterns) {
        $matchList = [regex]::Matches($pdfText, $pat, "IgnoreCase")
        foreach ($m in $matchList) {
            # Take only the first line of the capture
            $snippet = ($m.Groups[1].Value -split "[\r\n]")[0].Trim()
            if ($snippet.Length -gt 50) { $snippet = $snippet.Substring(0, 50) }

            # --- Try: ISO YYYY-MM-DD ---
            $iso = [regex]::Match($snippet, "(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})")
            if ($iso.Success) {
                $mm = $iso.Groups[2].Value.PadLeft(2,'0')
                $dd = $iso.Groups[3].Value.PadLeft(2,'0')
                $yy = $iso.Groups[1].Value
                return "$mm-$dd-$yy"
            }

            # --- Try: text month dates ---
            # DD MonthName YYYY  e.g. "29 August 2022"
            $t1 = [regex]::Match($snippet, "(\d{1,2})\s+([A-Za-z]{3,12})\.?\s*,?\s*(\d{4})")
            if ($t1.Success) {
                $day = $t1.Groups[1].Value.PadLeft(2,'0')
                $mn  = Get-MonthNumber $t1.Groups[2].Value
                $yr  = $t1.Groups[3].Value
                if ($mn -and $yr) { return "$mn-$day-$yr" }
            }

            # MonthName DD, YYYY  e.g. "August 29, 2022"
            $t2 = [regex]::Match($snippet, "([A-Za-z]{3,12})\.?\s+(\d{1,2}),?\s*(\d{4})")
            if ($t2.Success) {
                $mn  = Get-MonthNumber $t2.Groups[1].Value
                $day = $t2.Groups[2].Value.PadLeft(2,'0')
                $yr  = $t2.Groups[3].Value
                if ($mn -and $yr) { return "$mn-$day-$yr" }
            }

            # --- Try: numeric MM/DD/YYYY or DD.MM.YYYY ---
            $n = [regex]::Match($snippet, "(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})")
            if ($n.Success) {
                $p1 = $n.Groups[1].Value.PadLeft(2,'0')
                $p2 = $n.Groups[2].Value.PadLeft(2,'0')
                $p3 = $n.Groups[3].Value
                if ($p3.Length -eq 2) { $p3 = "20$p3" }
                return "$p1-$p2-$p3"
            }
        }
    }
    return $null
}

# =============================================================================
# FUNCTION: Create a styled .xlsx report using Excel COM
# =============================================================================
function Export-ToExcel {
    param(
        [array]  $Data,
        [array]  $Headers,
        [string] $FilePath,
        [string] $SheetName
    )
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = $SheetName

    # Header row
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $cell = $ws.Cells.Item(1, $c + 1)
        $cell.Value2 = $Headers[$c]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 8388608    # Dark blue in Excel BGR
        $cell.Font.Color = 16777215       # White
    }

    # Data rows
    $row = 2
    foreach ($record in $Data) {
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $ws.Cells.Item($row, $c + 1).Value2 = [string]$record[$Headers[$c]]
        }
        $row++
    }

    $ws.UsedRange.Columns.AutoFit() | Out-Null
    $wb.SaveAs($FilePath, 51)   # 51 = xlOpenXMLWorkbook
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}

# =============================================================================
# READ EXCEL SHEET: Extract Document IDs, Product Names, and URLs
# =============================================================================
Write-Host ""
Write-Host "[1/4] Reading Excel file..." -ForegroundColor Cyan
Copy-Item -Path $excelFile -Destination $tempCopy -Force
if (Test-Path $extractPath) { Remove-Item -Recurse -Force $extractPath }
[System.IO.Compression.ZipFile]::ExtractToDirectory($tempCopy, $extractPath)

$sharedStrings = @()
if (Test-Path "$extractPath\xl\sharedStrings.xml") {
    [xml]$ssXml = Get-Content "$extractPath\xl\sharedStrings.xml"
    foreach ($si in $ssXml.sst.si) {
        if ($si.t)      { $sharedStrings += $si.t }
        elseif ($si.r)  { $sharedStrings += ($si.r | ForEach-Object { $_.t }) -join "" }
        else            { $sharedStrings += "" }
    }
}

$hyperlinks = @{}
if (Test-Path "$extractPath\xl\worksheets\_rels\sheet1.xml.rels") {
    [xml]$relsXml = Get-Content "$extractPath\xl\worksheets\_rels\sheet1.xml.rels"
    foreach ($rel in $relsXml.Relationships.Relationship) {
        if ($rel.TargetMode -eq "External") { $hyperlinks[$rel.Id] = $rel.Target }
    }
}

[xml]$sheetXml = Get-Content "$extractPath\xl\worksheets\sheet1.xml"

$cellHyperlinks = @{}
if ($sheetXml.worksheet.hyperlinks) {
    foreach ($hl in $sheetXml.worksheet.hyperlinks.hyperlink) {
        $rId = $hl.getAttribute("r:id")
        if (-not $rId) { $rId = $hl.'id' }
        if ($hyperlinks.ContainsKey($rId)) { $cellHyperlinks[$hl.ref] = $hyperlinks[$rId] }
    }
}

$items = @()
foreach ($row in $sheetXml.worksheet.sheetData.row) {
    $rowNum = [int]$row.r
    if ($rowNum -eq 1) { continue }

    $cells = @{}
    foreach ($c in $row.c) {
        $col = $c.r -replace '[0-9]',''
        $val = $c.v
        if ($c.t -eq "s" -and $null -ne $val) { $val = $sharedStrings[[int]$val] }
        $cells[$col] = $val
    }

    $docId       = $cells["A"]
    $productName = $cells["B"]
    $url = $cellHyperlinks["M$rowNum"]
    if (-not $url) { $url = $cells["M"] }
    if (-not $url) { $url = $cellHyperlinks["W$rowNum"] }
    if (-not $url) { $url = $cells["W"] }

    if (-not $docId -or -not $url) { continue }

    $items += [PSCustomObject]@{
        Row         = $rowNum
        DocID       = $docId
        ProductName = $productName
        URL         = $url
    }
}

Write-Host "   Found $($items.Count) records in Excel." -ForegroundColor Green
if ($Limit -gt 0) { $items = $items | Select-Object -First $Limit }

# =============================================================================
# PROCESS ALL FILES
# =============================================================================
Write-Host "[2/4] Processing $($items.Count) files..." -ForegroundColor Cyan

$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

$downloadReport   = @()
$noRevisionReport = @()
$successCount     = 0
$failCount        = 0
$noRevCount       = 0

foreach ($item in $items) {
    $tempFile  = Join-Path $tempDownload "$($item.DocID)_temp.pdf"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    Write-Host "Row $($item.Row) [DocID: $($item.DocID)] " -NoNewline

    try {
        # --- Step 1: Download to temp folder ---
        $webClient.DownloadFile($item.URL, $tempFile)

        # --- Step 2: Extract text from PDF ---
        $pdfText = Extract-PdfText $tempFile

        # --- Step 3: Parse revision date (multi-language) ---
        $sdsDate = Parse-SdsRevisionDate $pdfText

        if ($sdsDate) {
            # HAS revision date -> save to Renewals
            $finalFile = "$($item.DocID)_${sdsDate}.pdf"
            $destPath  = Join-Path $outputFolder $finalFile
            Move-Item -Path $tempFile -Destination $destPath -Force

            Write-Host "[SUCCESS] $sdsDate -> $finalFile" -ForegroundColor Green

            $downloadReport += [ordered]@{
                "Document ID"   = $item.DocID
                "Product Name"  = $item.ProductName
                "URL"           = $item.URL
                "Status"        = "SUCCESS"
                "Revision Date" = $sdsDate
                "Saved As"      = $finalFile
                "Folder"        = "Renewals"
                "Timestamp"     = $timestamp
                "Error"         = ""
            }
            $successCount++

        } else {
            # NO revision date -> save to no_revisions
            $finalFile = "$($item.DocID).pdf"
            $destPath  = Join-Path $noRevFolder $finalFile
            Move-Item -Path $tempFile -Destination $destPath -Force

            Write-Host "[NO DATE] No revision date found -> $finalFile" -ForegroundColor Yellow

            $downloadReport += [ordered]@{
                "Document ID"   = $item.DocID
                "Product Name"  = $item.ProductName
                "URL"           = $item.URL
                "Status"        = "NO REVISION DATE"
                "Revision Date" = "Not Found"
                "Saved As"      = $finalFile
                "Folder"        = "Renewals\no_revisions"
                "Timestamp"     = $timestamp
                "Error"         = ""
            }
            $noRevisionReport += [ordered]@{
                "Document ID"  = $item.DocID
                "Product Name" = $item.ProductName
                "URL"          = $item.URL
                "Saved As"     = $finalFile
                "Timestamp"    = $timestamp
            }
            $noRevCount++
        }

    } catch {
        $errMsg = $_.ToString()
        Write-Host "[FAILED] $errMsg" -ForegroundColor Red
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

        $downloadReport += [ordered]@{
            "Document ID"   = $item.DocID
            "Product Name"  = $item.ProductName
            "URL"           = $item.URL
            "Status"        = "FAILED"
            "Revision Date" = ""
            "Saved As"      = ""
            "Folder"        = ""
            "Timestamp"     = $timestamp
            "Error"         = $errMsg
        }
        $failCount++
    }
}

# =============================================================================
# GENERATE EXCEL REPORTS
# =============================================================================
Write-Host ""
Write-Host "[3/4] Generating Excel reports..." -ForegroundColor Cyan

$reportHeaders  = @("Document ID","Product Name","URL","Status","Revision Date","Saved As","Folder","Timestamp","Error")
$reportPath     = "$baseDir\download_report.xlsx"

Export-ToExcel -Data $downloadReport -Headers $reportHeaders -FilePath $reportPath -SheetName "Download Report"
Write-Host "   Created: $reportPath" -ForegroundColor Green

if ($noRevisionReport.Count -gt 0) {
    $noRevHeaders    = @("Document ID","Product Name","URL","Saved As","Timestamp")
    $noRevReportPath = "$noRevFolder\no_revisions_report.xlsx"
    Export-ToExcel -Data $noRevisionReport -Headers $noRevHeaders -FilePath $noRevReportPath -SheetName "No Revision Date"
    Write-Host "   Created: $noRevReportPath" -ForegroundColor Green
}

# =============================================================================
# SUMMARY
# =============================================================================
Write-Host ""
Write-Host "[4/4] === SUMMARY ===================================================" -ForegroundColor Cyan
Write-Host "   Total processed        : $($items.Count)"
Write-Host "   Saved with date        : $successCount  ->  $outputFolder"
Write-Host "   No revision date found : $noRevCount   ->  $noRevFolder"
Write-Host "   Failed to download     : $failCount"
Write-Host "   Main report            : $reportPath"
if ($noRevisionReport.Count -gt 0) {
    Write-Host "   No-Rev report          : $noRevFolder\no_revisions_report.xlsx"
}
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""
