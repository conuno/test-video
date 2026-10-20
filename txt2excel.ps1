<#
.SYNOPSIS
    Converts every report txtResults\<name>.txt (written by videoqa.ps1) into
    excelResults\<name>.xlsx. Existing .xlsx files are overwritten.

    The .xlsx file is written directly (Office Open XML = zip of XML files),
    so neither Excel nor any PowerShell module is needed.

    Sheet "Steps":            Abschnitt | Nr. | Action | Input | Testdaten / Variablen | Expected result
    Sheet "Voraussetzungen":  Abschnitt | Voraussetzung
    Sections without a steps table (e.g. "KEINE SICHTBARE AENDERUNG" or an error text)
    become one row whose Action column holds that text.

.PARAMETER SourceDir
    Folder with the .txt reports. Default: txtResults next to this script.
.PARAMETER TargetDir
    Folder for the .xlsx files. Default: excelResults next to this script.
.PARAMETER File
    Convert only these report files instead of the whole SourceDir
    (videoqa.ps1 uses this right after a report is finished).
#>
[CmdletBinding()]
param(
    [string]$SourceDir = '',
    [string]$TargetDir = '',
    [string[]]$File = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

if (-not $SourceDir) { $SourceDir = Join-Path $PSScriptRoot 'txtResults' }
if (-not $TargetDir) { $TargetDir = Join-Path $PSScriptRoot 'excelResults' }
$LogDir  = Join-Path $PSScriptRoot 'logs'
$LogFile = Join-Path $LogDir 'txt2excel.log'

$Utf8NoBom = New-Object System.Text.UTF8Encoding $false
$Utf8Bom   = New-Object System.Text.UTF8Encoding $true
$Invariant = [Globalization.CultureInfo]::InvariantCulture
$MaxCellLength = 32767   # Excel limit per cell

$StepHeaders = @('Abschnitt', 'Nr.', 'Action', 'Input', 'Testdaten / Variablen', 'Expected result')
$StepWidths  = @(11, 6, 32, 50, 40, 60)
$PrereqHeaders = @('Abschnitt', 'Voraussetzung')
$PrereqWidths  = @(11, 120)

# --------------------------------------------------------------- helpers

function Write-Log {
    param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss', $Invariant), $Level, $Message
    Write-Host $line
    try { [IO.File]::AppendAllText($LogFile, $line + "`r`n", $Utf8Bom) } catch { }
}

# Markdown cell text -> Excel cell text: <br> becomes a line break, **bold** markers go.
function Convert-CellText {
    param([AllowEmptyString()][string]$Text)
    $t = [regex]::Replace($Text, '(?i)<br\s*/?>', "`n")
    $t = $t.Replace('**', '')
    return (($t -split "`n") | ForEach-Object { $_.Trim() }) -join "`n"
}

# "| a | b \| c | d |" -> 'a', 'b | c', 'd'
function Split-TableRow {
    param([string]$Line)
    $inner = $Line.Trim()
    if ($inner.StartsWith('|')) { $inner = $inner.Substring(1) }
    if ($inner.EndsWith('|') -and -not $inner.EndsWith('\|')) { $inner = $inner.Substring(0, $inner.Length - 1) }
    return , @([regex]::Split($inner, '(?<!\\)\|') | ForEach-Object { $_.Replace('\|', '|').Trim() })
}

# Whole number -> [int], anything else stays text.
function ConvertTo-CellValue {
    param([AllowEmptyString()][string]$Text)
    $n = 0
    if ($Text -match '^\d{1,9}$' -and [int]::TryParse($Text, [ref]$n)) { return $n }
    return $Text
}

# Parses a videoqa report into step rows and prerequisite rows.
function ConvertFrom-Report {
    param([string]$Text)

    $steps   = New-Object System.Collections.Generic.List[object]
    $prereqs = New-Object System.Collections.Generic.List[object]
    $content = $Text -replace "`r`n", "`n"
    $sections = [regex]::Matches($content, '(?m)^--- Abschnitt (\d+) ---[ \t]*$')

    for ($i = 0; $i -lt $sections.Count; $i++) {
        $number = [int]$sections[$i].Groups[1].Value
        $start  = $sections[$i].Index + $sections[$i].Length
        $end    = if ($i + 1 -lt $sections.Count) { $sections[$i + 1].Index } else { $content.Length }
        $body   = $content.Substring($start, $end - $start).Trim()

        $heading = ''
        $stepCountBefore = $steps.Count
        $hasPrereqs = $false
        foreach ($line in ($body -split "`n")) {
            $t = $line.Trim()
            if ($t -match '^##\s+(.+)$') { $heading = $Matches[1].Trim(); continue }

            if ($heading -eq 'Voraussetzungen' -and $t -match '^[-*]\s+(.+)$') {
                $prereqs.Add(@($number, (Convert-CellText $Matches[1])))
                $hasPrereqs = $true
                continue
            }

            if ($t.StartsWith('|')) {
                $cells = Split-TableRow $t
                if ($cells.Count -eq 0) { continue }
                if (@($cells | Where-Object { $_ -notmatch '^:?-{2,}:?$' }).Count -eq 0) { continue }   # |---|---|
                if ($cells[0] -eq '#') { continue }                                                     # header row
                if ($cells.Count -gt 5) { $cells = @($cells[0..3]) + @(($cells[4..($cells.Count - 1)]) -join ' | ') }
                while ($cells.Count -lt 5) { $cells += '' }
                $steps.Add(@(
                    $number,
                    (ConvertTo-CellValue $cells[0]),
                    (Convert-CellText $cells[1]),
                    (Convert-CellText $cells[2]),
                    (Convert-CellText $cells[3]),
                    (Convert-CellText $cells[4])
                ))
            }
        }

        if ($steps.Count -eq $stepCountBefore) {
            # No steps table: keep the section text visible as one row.
            $note = if ($hasPrereqs) { '(keine Testschritte)' } else { Convert-CellText $body }
            $steps.Add(@($number, '', $note, '', '', ''))
        }
    }
    return [pscustomobject]@{ Sections = $sections.Count; Steps = $steps.ToArray(); Prereqs = $prereqs.ToArray() }
}

# ----------------------------------------------------------- xlsx writer

function ConvertTo-XmlText {
    param([string]$Text)
    # Remove characters that are not allowed in XML 1.0, cap at Excel's cell limit.
    $t = [regex]::Replace($Text, ('[\x00-\x08\x0B\x0C\x0E-\x1F' + [char]0xFFFE + [char]0xFFFF + ']'), '')
    if ($t.Length -gt $MaxCellLength) { $t = $t.Substring(0, $MaxCellLength) }
    return $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Get-ColumnName {
    param([int]$Index)   # 1-based
    $name = ''
    while ($Index -gt 0) {
        $rem = ($Index - 1) % 26
        $name = [char](65 + $rem) + $name
        $Index = [int][Math]::Floor(($Index - 1) / 26)
    }
    return $name
}

# Style ids from styles.xml: 1 = wrap + top, 2 = bold + wrap + top (header)
function Get-SheetXml {
    param([string[]]$Headers, [int[]]$Widths, [object[]]$Rows)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<cols>')
    for ($c = 0; $c -lt $Widths.Count; $c++) {
        [void]$sb.AppendFormat($Invariant, '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($c + 1), $Widths[$c])
    }
    [void]$sb.Append('</cols><sheetData>')

    $allRows = @(, $Headers) + $Rows
    for ($r = 0; $r -lt $allRows.Count; $r++) {
        $rowNumber = $r + 1
        $style = if ($r -eq 0) { 2 } else { 1 }
        [void]$sb.AppendFormat($Invariant, '<row r="{0}">', $rowNumber)
        $values = @($allRows[$r])
        for ($c = 0; $c -lt $values.Count; $c++) {
            $ref = '{0}{1}' -f (Get-ColumnName ($c + 1)), $rowNumber
            $value = $values[$c]
            if ($value -is [int]) {
                [void]$sb.AppendFormat($Invariant, '<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $style, $value)
            } elseif ([string]$value -ne '') {
                [void]$sb.AppendFormat($Invariant, '<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>',
                    $ref, $style, (ConvertTo-XmlText ([string]$value)))
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.AppendFormat($Invariant, '<autoFilter ref="A1:{0}{1}"/>', (Get-ColumnName $Headers.Count), $allRows.Count)
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function Add-ZipEntry {
    param($Zip, [string]$Name, [string]$Text)
    $entry  = $Zip.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $Utf8NoBom)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

# $Sheets: array of @{ Name; Headers; Widths; Rows }
function Write-Xlsx {
    param([string]$Path, [object[]]$Sheets)

    $ns    = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $pkgNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $head  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

    $types = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $sheetList = ''; $sheetRels = ''; $filters = ''
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $n = $i + 1
        $sheet = $Sheets[$i]
        $types     += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $sheetList += "<sheet name=`"$($sheet.Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
        $sheetRels += "<Relationship Id=`"rId$n`" Type=`"$relNs/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        $lastRow = @($sheet.Rows).Count + 1
        $filters += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$i`" hidden=`"1`">'$($sheet.Name)'!`$A`$1:`$$(Get-ColumnName $sheet.Headers.Count)`$$lastRow</definedName>"
    }
    $types += '</Types>'
    $rootRels = $head + "<Relationships xmlns=`"$pkgNs`"><Relationship Id=`"rId1`" Type=`"$relNs/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>"
    $workbook = $head + "<workbook xmlns=`"$ns`" xmlns:r=`"$relNs`"><sheets>$sheetList</sheets><definedNames>$filters</definedNames></workbook>"
    $stylesId = $Sheets.Count + 1
    $workbookRels = $head + "<Relationships xmlns=`"$pkgNs`">$sheetRels<Relationship Id=`"rId$stylesId`" Type=`"$relNs/styles`" Target=`"styles.xml`"/></Relationships>"
    $styles = $head + "<styleSheet xmlns=`"$ns`">" +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

    # Write to a temp file first, so an existing .xlsx is only replaced by a complete one.
    $tmp = "$Path.tmp"
    $stream = [IO.File]::Open($tmp, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            Add-ZipEntry $zip '[Content_Types].xml' $types
            Add-ZipEntry $zip '_rels/.rels' $rootRels
            Add-ZipEntry $zip 'xl/workbook.xml' $workbook
            Add-ZipEntry $zip 'xl/_rels/workbook.xml.rels' $workbookRels
            Add-ZipEntry $zip 'xl/styles.xml' $styles
            for ($i = 0; $i -lt $Sheets.Count; $i++) {
                Add-ZipEntry $zip "xl/worksheets/sheet$($i + 1).xml" (Get-SheetXml $Sheets[$i].Headers $Sheets[$i].Widths $Sheets[$i].Rows)
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
    try {
        Move-Item -LiteralPath $tmp -Destination $Path -Force
    } catch {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        throw "Cannot replace $Path (is it open in Excel?): $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------ main

foreach ($dir in @($TargetDir, $LogDir)) {
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}
$failed = 0
$converted = 0
if (@($File).Count -gt 0) {
    $files = @()
    foreach ($path in $File) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $files += Get-Item -LiteralPath $path
        } else {
            Write-Log "Report not found: $path" 'ERROR'
            $failed++
        }
    }
    Write-Log "Converting $($files.Count) report(s) to $TargetDir"
} else {
    if (-not (Test-Path -LiteralPath $SourceDir -PathType Container)) {
        Write-Log "Source folder not found: $SourceDir" 'ERROR'
        exit 1
    }
    # Only finished reports: *.txt (not *.txt.partial).
    $files = @(Get-ChildItem -LiteralPath $SourceDir -File | Where-Object { $_.Extension -eq '.txt' } | Sort-Object Name)
    Write-Log "Converting $($files.Count) report(s) from $SourceDir to $TargetDir"
}

foreach ($txtFile in $files) {
    $target = Join-Path $TargetDir ($txtFile.BaseName + '.xlsx')
    try {
        $report = ConvertFrom-Report ([IO.File]::ReadAllText($txtFile.FullName, [Text.Encoding]::UTF8))
        if ($report.Sections -eq 0) { Write-Log "$($txtFile.Name): no '--- Abschnitt N ---' sections found" 'WARN' }
        Write-Xlsx -Path $target -Sheets @(
            @{ Name = 'Steps';           Headers = $StepHeaders;   Widths = $StepWidths;   Rows = $report.Steps },
            @{ Name = 'Voraussetzungen'; Headers = $PrereqHeaders; Widths = $PrereqWidths; Rows = $report.Prereqs }
        )
        Write-Log ("{0} -> {1} ({2} sections, {3} step rows, {4} prerequisites)" -f
            $txtFile.Name, $target, $report.Sections, @($report.Steps).Count, @($report.Prereqs).Count)
        $converted++
    } catch {
        $failed++
        Write-Log "$($txtFile.Name): $($_.Exception.Message)" 'ERROR'
    }
}
Write-Log "Done: $converted converted, $failed failed"
if ($failed -gt 0) { exit 1 }
exit 0
