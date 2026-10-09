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
# CONFIGURATION & PATHS
# =============================================================================
$baseDir      = "c:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF"
$excelFile    = "$baseDir\Book1.xlsx"
$outputFolder = "$baseDir\Renewals"
$noRevFolder  = "$baseDir\Renewals\no_revisions"
$tempDownload = "$env:TEMP\sds_temp_downloads"
$tempCopy     = "$env:TEMP\temp_Book1.xlsx"
$extractPath  = "$env:TEMP\excel_extracted"

# Excel Column Mappings
$colDocId   = "A"
$colProduct = "B"
$colUrl1    = "M"
$colUrl2    = "W"

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
# =============================================================================
# MULTI-LANGUAGE REVISION DATE KEYWORDS (STRICT REVISION DATE ONLY)
# =============================================================================
$revisionKeywords = @(
    # English
    "Date\s*of\s*issue\s*[\/\\]\s*Date\s*of\s*revision",
    "Issue\s*Date\s*[\/\\]\s*Revision\s*Date",
    "Date\s*of\s*(?:Last\s*)?Revision",
    "Revision\s*Date",     "Date\s*of\s*Revision",
    "Revised\s*Date",      "Date\s*Revised",
    "Revised\s*on",        "Revised",
    "Last\s*Revision(?:\s*Date)?", "Last\s*Revised",
    "Date\s*of\s*(?:Last\s*)?Update", "Update\s*Date",
    "Updated\s*on",        "Updated", "Last\s*Updated",
    "Rev\.\s*Date",        "Rev\s*Date",
    "Rev\s*[:\.\-]",       "Rev\.\s*[:\.\-]",
    "Revision\s*[:\.\-]",  "Revisiondate", "Revision",

    # German
    "Ausgabedatum\s*[\/\\]\s*[ÜüUeue]+berarbeitungsdatum",
    "Datum\s*der\s*(?:letzten\s*)?[ÜüUeue]+berarbeitung",
    "[ÜüUeue]+berarbeitungsdatum",
    "[ÜüUeue]+berarbeitet\s*am",
    "[ÜüUeue]+berarbeitet",
    "Revisionsdatum",      "Datum\s*der\s*Revision",
    "Revision\s*vom",      "Revidiert\s*am",
    "Stand\s*der\s*Information", "Stand\s*vom",

    # French
    "Date\s*d['’]?[eE]mission\s*[\/\\]\s*Date\s*de\s*r[eéèEÉÈ]vision",
    "Date\s*de\s*(?:la\s*)?(?:derni[eèéEÈÉ]re\s*)?r[eéèEÉÈ]vision",
    "Date\s*de\s*(?:derni[eèéEÈÉ]re\s*)?mise\s*(?:[aAàÀ]\s*)?jour",
    "Mise\s*(?:[aAàÀ]\s*)?jour\s*(?:le)?",
    "R[eéèEÉÈ]vis[eéèEÉÈ]\s*le",
    "Derni[eèéEÈÉ]re\s*r[eéèEÉÈ]vision",

    # Spanish
    "Fecha\s*de\s*emisi[oóOÓ]n\s*[\/\\]\s*Fecha\s*de\s*revisi[oóOÓ]n",
    "Fecha\s*de\s*(?:la\s*)?(?:[uúUÚ]ltima\s*)?revisi[oóOÓ]n",
    "Fecha\s*de\s*(?:[uúUÚ]ltima\s*)?actualizaci[oóOÓ]n",
    "Actualizado\s*(?:el)?", "Revisado\s*(?:el)?",
    "[UúUÚ]ltima\s*revisi[oóOÓ]n",

    # Italian
    "Data\s*di\s*emissione\s*[\/\\]\s*Data\s*di\s*revisione",
    "Data\s*di\s*(?:dell['’]ultima\s*)?revisione",
    "Data\s*di\s*(?:dell['’]ultimo\s*)?aggiornamento",
    "Aggiornato\s*il",     "Revisionato\s*il",

    # Portuguese
    "Data\s*de\s*emiss[aãAÃ]o\s*[\/\\]\s*Data\s*de\s*revis[aãAÃ]o",
    "Data\s*de\s*(?:da\s*)?(?:[uúUÚ]ltima\s*)?revis[aãAÃ]o",
    "Data\s*de\s*(?:da\s*)?(?:[uúUÚ]ltima\s*)?atualiza[cçCÇ][aãAÃ]o",
    "Revisado\s*em",       "Atualizado\s*em",

    # Dutch
    "Herzieningsdatum",    "Datum\s*van\s*herziening",
    "Revisiedatum",        "Datum\s*van\s*bijwerking",
    "Bijgewerkt\s*op",     "Herzien\s*op"
)

function Get-Patterns($keywordList) {
    $combinedKw = "(?:" + ($keywordList -join "|") + ")"
    return @(
        "$combinedKw\s*[:\-]?\s*([A-Za-z0-9\s,\.\/\-]+)",
        "$combinedKw[\s\r\n]+([A-Za-z0-9\s,\.\/\-]+)"
    )
}
$revisionPatterns = Get-Patterns $revisionKeywords

# =============================================================================
# FUNCTION: Parse Revision Date from SDS text (multi-language)
# =============================================================================
function Parse-SdsRevisionDate([string]$pdfText) {
    if (-not $pdfText) { return $null }

    foreach ($pat in $revisionPatterns) {
        $matchList = [regex]::Matches($pdfText, $pat, "IgnoreCase")
        foreach ($m in $matchList) {
            $snippet = ($m.Groups[1].Value -split "[\r\n]")[0].Trim()
            $snippet = ($snippet -split "(?i)(?:Date\s*of\s*previous|Datum\s*der\s*letzten|Previous\s*issue|Ausgabe|Druckdatum|Version\s*:)")[0].Trim()
            if ($snippet.Length -gt 60) { $snippet = $snippet.Substring(0, 60) }

            $iso = [regex]::Match($snippet, "(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})")
            if ($iso.Success) {
                $mm = $iso.Groups[2].Value.PadLeft(2,'0')
                $dd = $iso.Groups[3].Value.PadLeft(2,'0')
                $yy = $iso.Groups[1].Value
                return "$mm-$dd-$yy"
            }

            $neu = [regex]::Match($snippet, "(\d{1,2})\.(\d{1,2})\.(\d{4})")
            if ($neu.Success) {
                $dd = $neu.Groups[1].Value.PadLeft(2,'0')
                $mm = $neu.Groups[2].Value.PadLeft(2,'0')
                $yy = $neu.Groups[3].Value
                return "$mm-$dd-$yy"
            }

            $t1 = [regex]::Match($snippet, "(\d{1,2})[\s\-\.\/]+([A-Za-z]{3,12})\.?[\s\-\.\/,]*(\d{4})")
            if ($t1.Success) {
                $day = $t1.Groups[1].Value.PadLeft(2,'0')
                $mn  = Get-MonthNumber $t1.Groups[2].Value
                $yr  = $t1.Groups[3].Value
                if ($mn -and $yr) { return "$mn-$day-$yr" }
            }

            $t2 = [regex]::Match($snippet, "([A-Za-z]{3,12})\.?[\s\-\.\/]+(\d{1,2})[\s\-\.\/,]*(\d{4})")
            if ($t2.Success) {
                $mn  = Get-MonthNumber $t2.Groups[1].Value
                $day = $t2.Groups[2].Value.PadLeft(2,'0')
                $yr  = $t2.Groups[3].Value
                if ($mn -and $yr) { return "$mn-$day-$yr" }
            }

            $n = [regex]::Match($snippet, "(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})")
            if ($n.Success) {
                $p1 = [int]$n.Groups[1].Value
                $p2 = [int]$n.Groups[2].Value
                $p3 = $n.Groups[3].Value
                if ($p3.Length -eq 2) { $p3 = "20$p3" }
                if ($p1 -gt 12 -and $p2 -le 12) {
                    return ("{0:D2}-{1:D2}-$p3" -f $p2, $p1)
                } else {
                    return ("{0:D2}-{1:D2}-$p3" -f $p1, $p2)
                }
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

    $docId       = $cells[$colDocId]
    $productName = $cells[$colProduct]
    $url = $cellHyperlinks["${colUrl1}$rowNum"]
    if (-not $url) { $url = $cells[$colUrl1] }
    if (-not $url) { $url = $cellHyperlinks["${colUrl2}$rowNum"] }
    if (-not $url) { $url = $cells[$colUrl2] }

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
