<#
.SYNOPSIS
    Video QA analysis: extracts frames with ffmpeg and analyses them window by
    window with headless Claude Code. Windows PowerShell 5.1 port of the n8n
    workflow "Videoanalyse_OpenAI" (Request HV Format path).
    Each video gives one report txtResults\<name>.txt. Videos are not moved;
    a video whose report already exists is skipped.

.PARAMETER Video
    One or more video files and/or folders (a folder means all videos in it,
    sorted by name). Processed one after another in the given order.
.PARAMETER TestClaude
    One Claude Code call with the instruction "Antworte nur mit: OK".
.PARAMETER DryRun
    Extract frames and write all auftrag_*.txt files; no Claude calls, keep work\.
.PARAMETER TestOneWindow
    Analyse only window 1 of the first video and print the answer; keep work\.
.PARAMETER KeepWork
    Do not delete work\<name>\ after processing.

.EXAMPLE
    run.bat "D:\videos"
.EXAMPLE
    run.bat "D:\a.mp4" "D:\b.mp4" -KeepWork
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Video = @(),
    [switch]$TestClaude,
    [switch]$DryRun,
    [switch]$TestOneWindow,
    [switch]$KeepWork
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================ Configuration ============================
$BaseDir            = $PSScriptRoot
$FfmpegPath         = Join-Path $BaseDir 'ffmpeg\ffmpeg.exe'
$VideoExtensions    = @('.mp4')
$Fps                = 2
$WindowSeconds      = 15
$ContextSeconds     = 1
$JpegQuality        = 3        # ffmpeg -q:v
$ScaleFilter        = ''       # e.g. 'scale=1920:-2'; empty = native resolution
$ClaudeModel        = 'claude-sonnet-5-5'   # empty = Claude Code default (Opus); 'claude-haiku-5-5' = cheapest
$ClaudeTimeoutMin   = 15       # per window
$ClaudeRetryDelays  = @(30, 60, 120)   # seconds; i.e. up to 3 retries
$MinFileAgeSeconds  = 60       # skip videos that may still be copying
$ConvertToExcel     = $true    # convert each finished report to excelResults\<name>.xlsx
# =======================================================================

$ResultDir  = Join-Path $BaseDir 'txtResults'   # one report <name>.txt per video
$LogDir     = Join-Path $BaseDir 'logs'         # videoqa.log and <name>.error.txt
$WorkDir    = Join-Path $BaseDir 'work'         # temporary, one folder per video
$PromptFile = Join-Path $BaseDir 'prompt.txt'
$LogFile    = Join-Path $LogDir 'videoqa.log'
$LockFile   = Join-Path $WorkDir 'videoqa.lock'

$Utf8NoBom  = New-Object System.Text.UTF8Encoding $false
$Utf8Bom    = New-Object System.Text.UTF8Encoding $true
$Invariant  = [Globalization.CultureInfo]::InvariantCulture

# Fixed instruction for each window (ASCII only, no quotes or <>). {0} = task file name.
$ClaudeInstruction = 'Lies die Datei {0} in diesem Ordner vollstaendig und befolge sie exakt. ' +
    'Lies alle dort aufgefuehrten Bilddateien in der angegebenen Reihenfolge. ' +
    'Gib ausschliesslich das geforderte Ergebnis aus.'
$FallbackAnswer   = 'FEHLER: Claude Code lieferte keine auswertbare Antwort.'
# "KEINE SICHTBARE AENDERUNG" with a real umlaut; built from a char code so this file stays ASCII.
$NoChangeAnswer   = 'KEINE SICHTBARE ' + [char]0x00C4 + 'NDERUNG'

# Last stderr of ffmpeg or claude, for the .error.txt file.
$script:LastToolStderr = ''
$script:VideosDone     = 0
$script:VideosFailed   = 0
$script:VideosSkipped  = 0

# --------------------------------------------------------------- helpers

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss', $Invariant), $Level, $Message
    Write-Host $line
    try {
        [IO.File]::AppendAllText($LogFile, $line + "`r`n", $Utf8Bom)
    } catch {
        Write-Host "(could not write log file: $($_.Exception.Message))"
    }
}

# JavaScript Math.round: halves round towards +infinity.
# [Math]::Round would use banker's rounding and differ from n8n.
function Get-JsRound {
    param([double]$Value)
    return [Math]::Floor($Value + 0.5)
}

# Same as n8n formatTime(): MM:SS.mmm, minutes are not wrapped into hours.
function Format-VideoTime {
    param([double]$Seconds)
    $totalMs = [long][Math]::Max(0, (Get-JsRound ($Seconds * 1000)))
    $mins = [long][Math]::Floor($totalMs / 60000)
    $secs = [long][Math]::Floor(($totalMs % 60000) / 1000)
    $ms   = $totalMs % 1000
    return '{0:00}:{1:00}.{2:000}' -f $mins, $secs, $ms
}

function Initialize-Folders {
    foreach ($dir in @($ResultDir, $LogDir, $WorkDir)) {
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }
}

# Exclusive lock file. DeleteOnClose removes it even if the process is killed.
# A lock file that nobody holds open is a leftover (e.g. power loss) and is removed.
function Enter-RunLock {
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            return [IO.FileStream]::new($LockFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite,
                [IO.FileShare]::None, 1, [IO.FileOptions]::DeleteOnClose)
        } catch {
            if ($attempt -eq 2) { return $null }
            try {
                $probe = [IO.File]::Open($LockFile, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                $probe.Close()
                Remove-Item -LiteralPath $LockFile -Force
                Write-Log 'Removed stale lock file from an earlier run' 'WARN'
            } catch {
                return $null
            }
        }
    }
    return $null
}

function Exit-RunLock {
    param($LockStream)
    if ($LockStream) { $LockStream.Dispose() }
    if (Test-Path -LiteralPath $LockFile) {
        try { Remove-Item -LiteralPath $LockFile -Force } catch { }
    }
}

# Get-Command with -CommandType Application returns claude.exe or claude.cmd,
# never the npm claude.ps1 shim that Start-Process cannot run.
# Falls back to the native installer's location, because Task Scheduler may start
# with a PATH from before ~\.local\bin was added.
function Resolve-Claude {
    $cmd = Get-Command claude -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) { return $cmd.Source }
    $native = Join-Path $env:USERPROFILE '.local\bin\claude.exe'
    if (Test-Path -LiteralPath $native) { return $native }
    return $null
}

# Returns the videos that are ready, natural-sorted by name.
# Expands the -Video arguments (files and/or folders) into the list of videos to
# process, in the given order; folders are natural-sorted by name. Duplicates,
# missing paths, non-videos and videos that may still be copying are left out.
function Get-InputVideos {
    param([string[]]$Paths)

    $extensions = @($VideoExtensions | ForEach-Object { $_.ToLowerInvariant() })
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $ready = New-Object System.Collections.Generic.List[object]

    foreach ($raw in $Paths) {
        # cmd turns "D:\videos\" into D:\videos" (the \" escapes the quote), so strip quotes.
        $path = $raw.Trim().Trim('"')
        if (-not $path) { continue }

        if (Test-Path -LiteralPath $path -PathType Container) {
            $candidates = @(Get-ChildItem -LiteralPath $path -File |
                Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
                Sort-Object { [regex]::Replace($_.Name, '\d+', { param($m) $m.Value.PadLeft(20, '0') }) })
            if ($candidates.Count -eq 0) {
                Write-Log "No videos ($($VideoExtensions -join ', ')) in folder $path" 'WARN'
            }
        } elseif (Test-Path -LiteralPath $path -PathType Leaf) {
            $item = Get-Item -LiteralPath $path
            if ($extensions -notcontains $item.Extension.ToLowerInvariant()) {
                Write-Log "Skipping '$path': not a video ($($VideoExtensions -join ', '))" 'WARN'
                continue
            }
            $candidates = @($item)
        } else {
            Write-Log "Input not found: $path" 'ERROR'
            continue
        }

        foreach ($file in $candidates) {
            if (-not $seen.Add($file.FullName)) { continue }
            if (Test-VideoReady $file) { $ready.Add($file) }
        }
    }
    return , $ready.ToArray()
}

# False (with a warning) if the video was changed very recently or is locked.
function Test-VideoReady {
    param($File)
    $age = ([DateTime]::Now - $File.LastWriteTime).TotalSeconds
    if ($age -lt $MinFileAgeSeconds) {
        Write-Log ([string]::Format($Invariant, "Skipping '{0}' for this run: modified {1:0} s ago", $File.FullName, $age)) 'WARN'
        return $false
    }
    try {
        $stream = [IO.File]::Open($File.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $stream.Close()
    } catch {
        Write-Log "Skipping '$($File.FullName)' for this run: file is locked (still copying?)" 'WARN'
        return $false
    }
    return $true
}

# Runs ffmpeg and returns the sorted absolute paths of the extracted frames.
function Invoke-FrameExtraction {
    param([string]$VideoPath, [string]$FrameDir)

    New-Item -ItemType Directory -Path $FrameDir -Force | Out-Null
    $vf = "fps=$Fps"
    if ($ScaleFilter) { $vf += ",$ScaleFilter" }
    $ffArgs = @(
        '-nostdin', '-hide_banner', '-loglevel', 'error', '-y',
        '-i', $VideoPath,
        '-vf', $vf,
        '-q:v', "$JpegQuality",
        (Join-Path $FrameDir 'frame_%06d.jpg')
    )

    # PowerShell 5.1 turns native stderr lines into ErrorRecords; with
    # ErrorActionPreference=Stop the first one would throw. Collect them instead.
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & $FfmpegPath @ffArgs 2>&1
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevEap
    }
    $stderr = (@($output) | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" }
    }) -join "`n"
    $script:LastToolStderr = $stderr

    if ($exitCode -ne 0) {
        throw "FFmpeg frame extraction failed (exit code $exitCode) for ${VideoPath}:`n$stderr"
    }

    $frames = @(Get-ChildItem -LiteralPath $FrameDir -File |
        Where-Object { $_.Name -match '^frame_\d+\.jpg$' } |
        Sort-Object Name |
        ForEach-Object { $_.FullName })
    if ($frames.Count -eq 0) {
        throw "FFmpeg produced no JPG frames in $FrameDir"
    }
    return , $frames
}

# Same windowing as the n8n node "Create Image Batches" (chunk offset is always 0).
function Get-Windows {
    param([string[]]$FramePaths)

    $frameCount    = $FramePaths.Count
    $primaryFrames = [int][Math]::Max(1, (Get-JsRound ($Fps * $WindowSeconds)))
    $contextFrames = [int][Math]::Max(0, (Get-JsRound ($Fps * $ContextSeconds)))

    $windows = New-Object System.Collections.Generic.List[object]
    $number = 1
    for ($primaryStart = 0; $primaryStart -lt $frameCount; $primaryStart += $primaryFrames) {
        $primaryEndExclusive = [Math]::Min($frameCount, $primaryStart + $primaryFrames)
        $includeStart = [Math]::Max(0, $primaryStart - $contextFrames)

        $frames = New-Object System.Collections.Generic.List[object]
        for ($i = $includeStart; $i -lt $primaryEndExclusive; $i++) {
            $frames.Add([pscustomobject]@{
                Path = $FramePaths[$i]
                Time = Format-VideoTime ($i / [double]$Fps)
                Role = $(if ($i -lt $primaryStart) { 'CONTEXT_ONLY' } else { 'ANALYZE' })
            })
        }

        $windows.Add([pscustomobject]@{
            Number       = $number
            PrimaryStart = Format-VideoTime ($primaryStart / [double]$Fps)
            PrimaryEnd   = Format-VideoTime (($primaryEndExclusive - 1) / [double]$Fps)
            Frames       = $frames.ToArray()
        })
        $number++
    }
    return , $windows.ToArray()
}

# Writes work\<name>\auftrag_NNNN.txt and returns its path.
function New-TaskFile {
    param($Window, [string]$VideoWorkDir, [string]$PromptTemplate)

    $text = $PromptTemplate.
        Replace('{{FPS}}', [string]$Fps).
        Replace('{{FRAME_INTERVAL_MS}}', [string](Get-JsRound (1000 / [double]$Fps))).
        Replace('{{PRIMARY_START}}', $Window.PrimaryStart).
        Replace('{{PRIMARY_END}}', $Window.PrimaryEnd)

    $lines = New-Object System.Collections.Generic.List[string]
    $index = 1
    foreach ($frame in $Window.Frames) {
        $lines.Add(('FRAME_IN_REQUEST={0:000} | ABSOLUTE_VIDEO_TIME={1} | LOCAL_CHUNK_TIME={1} | ROLE={2} | DATEI={3}' -f
            $index, $frame.Time, $frame.Role, $frame.Path))
        $index++
    }

    $content = $text.TrimEnd() + "`n`nBILDLISTE:`n" + ($lines -join "`n") + "`n"
    $path = Join-Path $VideoWorkDir ('auftrag_{0:0000}.txt' -f $Window.Number)
    [IO.File]::WriteAllText($path, $content, $Utf8NoBom)
    return $path
}

# Quotes one argument for the Windows command line (CommandLineToArgvW rules).
# Start-Process in PS 5.1 joins -ArgumentList without quoting, so we do it here.
function ConvertTo-CommandLineArgument {
    param([AllowEmptyString()][string]$Value)
    if ($Value -ne '' -and $Value -notmatch '[\s"]') { return $Value }
    $escaped = [regex]::Replace($Value, '(\\*)"', { param($m) ($m.Groups[1].Value * 2) + '\"' })
    $escaped = [regex]::Replace($escaped, '(\\+)$', { param($m) $m.Groups[1].Value * 2 })
    return '"' + $escaped + '"'
}

# UTF-8 file content, or '' if the file is missing or empty
# (Get-Content -Raw returns nothing at all for an empty file).
function Read-TextFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ($null -eq $content) { return '' }
    return [string]$content
}

# Property of a parsed JSON object, or $null if missing (StrictMode-safe).
function Get-JsonValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

# One headless Claude Code call. Never throws for Claude failures; returns Ok/Error instead.
#   --tools Read                only the Read tool exists in the session
#   --permission-mode dontAsk   anything not allowed by default is denied (Read outside
#                               the working folder is denied, Read inside it is allowed)
#   --permission-prompts none   nothing can ever wait for an interactive prompt
#   --setting-sources ""        ignore user/project/local settings (hooks, permissions)
#   --strict-mcp-config         no MCP servers
#   --no-session-persistence    no session files on disk
function Invoke-ClaudeOnce {
    param([string]$ClaudePath, [string]$WorkingDir, [string]$Instruction, [string]$Tag)

    $stdinFile  = Join-Path $WorkingDir 'claude_stdin.txt'
    $stdoutFile = Join-Path $WorkingDir "claude_$Tag.stdout.json"
    $stderrFile = Join-Path $WorkingDir "claude_$Tag.stderr.txt"
    [IO.File]::WriteAllText($stdinFile, '')

    $argv = @('-p', $Instruction, '--output-format', 'json',
        '--tools', 'Read', '--permission-mode', 'dontAsk', '--permission-prompts', 'none',
        '--setting-sources', '', '--strict-mcp-config', '--no-session-persistence')
    if ($ClaudeModel) { $argv += @('--model', $ClaudeModel) }
    $argString = (@($argv) | ForEach-Object { ConvertTo-CommandLineArgument $_ }) -join ' '

    $result = [pscustomobject]@{
        Ok = $false; Text = ''; Error = ''; Seconds = 0.0; Cost = $null; Turns = $null; Denials = 0; Json = $null
    }
    $started = [DateTime]::Now
    $proc = Start-Process -FilePath $ClaudePath -ArgumentList $argString -WorkingDirectory $WorkingDir `
        -NoNewWindow -PassThru -RedirectStandardInput $stdinFile `
        -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile
    $null = $proc.Handle   # PS 5.1: without this, ExitCode can be $null after WaitForExit

    $timedOut = -not $proc.WaitForExit($ClaudeTimeoutMin * 60000)
    if ($timedOut) {
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { & taskkill.exe /T /F /PID $proc.Id 2>&1 | Out-Null } finally { $ErrorActionPreference = $prevEap }
        $null = $proc.WaitForExit(30000)
    }
    $result.Seconds = ([DateTime]::Now - $started).TotalSeconds

    $script:LastToolStderr = (Read-TextFile $stderrFile).Trim()

    if ($timedOut) {
        $result.Error = "timeout after $ClaudeTimeoutMin min (process tree killed)"
        return $result
    }

    $stdout = Read-TextFile $stdoutFile
    $json = $null
    try {
        $json = $stdout | ConvertFrom-Json
    } catch {
        # Fall back to the last line that looks like a JSON object.
        $lastLine = @($stdout -split "`r?`n" | Where-Object { $_.TrimStart().StartsWith('{') }) | Select-Object -Last 1
        if ($lastLine) { try { $json = $lastLine | ConvertFrom-Json } catch { $json = $null } }
    }
    $result.Json    = $json
    $result.Cost    = Get-JsonValue $json 'total_cost_usd'
    $result.Turns   = Get-JsonValue $json 'num_turns'
    $result.Denials = @(Get-JsonValue $json 'permission_denials' | Where-Object { $_ }).Count
    $answer         = Get-JsonValue $json 'result'

    if ($proc.ExitCode -ne 0) {
        $result.Error = "exit code $($proc.ExitCode)"
        if ($answer) { $result.Error += ": $answer" } elseif ($script:LastToolStderr) { $result.Error += ": $($script:LastToolStderr)" }
    } elseif ($null -eq $json) {
        $result.Error = 'output is not valid JSON'
    } elseif ((Get-JsonValue $json 'is_error') -eq $true) {
        $result.Error = "is_error=true: $answer"
    } else {
        $result.Ok   = $true
        $result.Text = [string]$answer
    }
    return $result
}

# Calls Claude Code with retries ($ClaudeRetryDelays). Throws after the last failed attempt.
function Invoke-ClaudeAnalysis {
    param([string]$ClaudePath, [string]$WorkingDir, [string]$Instruction, [string]$Label, [string]$Tag)

    $maxAttempts = @($ClaudeRetryDelays).Count + 1
    $last = $null
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($attempt -gt 1) {
            $delay = @($ClaudeRetryDelays)[$attempt - 2]
            Write-Log "${Label}: waiting $delay s before attempt $attempt/$maxAttempts" 'WARN'
            Start-Sleep -Seconds $delay
        }
        $last = Invoke-ClaudeOnce -ClaudePath $ClaudePath -WorkingDir $WorkingDir -Instruction $Instruction `
            -Tag ('{0}_try{1}' -f $Tag, $attempt)
        $info = [string]::Format($Invariant, '{0:0} s, attempt {1}/{2}, turns {3}, cost {4:0.0000} USD',
            $last.Seconds, $attempt, $maxAttempts, $last.Turns, $last.Cost)
        if ($last.Ok) {
            Write-Log "${Label}: OK ($info)"
            if ($last.Denials -gt 0) { Write-Log "${Label}: Claude Code reported $($last.Denials) denied tool call(s)" 'WARN' }
            return $last
        }
        Write-Log "${Label}: attempt failed ($info): $($last.Error)" 'WARN'
    }
    throw "Claude Code failed for $Label after $maxAttempts attempts: $($last.Error)"
}

# Section 6.7: keep the answer from "## Voraussetzungen", or exactly the no-change marker.
function Get-CleanAnswer {
    param([AllowEmptyString()][string]$Text, [string]$Label)

    $trimmed = ($Text -replace "`r`n", "`n").Trim()
    if ($trimmed -eq '') {
        Write-Log "${Label}: empty answer, using fallback text" 'WARN'
        return $FallbackAnswer
    }
    $index = $trimmed.IndexOf('## Voraussetzungen', [StringComparison]::Ordinal)
    if ($index -ge 0) {
        if ($index -gt 0) { Write-Log "${Label}: removed $index characters before '## Voraussetzungen'" 'WARN' }
        return $trimmed.Substring($index).Trim()
    }
    $lines = @($trimmed -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    if ($lines -ccontains $NoChangeAnswer) {
        if ($lines.Count -gt 1) { Write-Log "${Label}: kept only '$NoChangeAnswer' from a longer answer" 'WARN' }
        return $NoChangeAnswer
    }
    Write-Log "${Label}: answer has neither '## Voraussetzungen' nor '$NoChangeAnswer'; keeping raw text" 'WARN'
    return $trimmed
}

# Same layout as the n8n "Write Batch Analysis" file: UTF-8 with BOM, \n line endings.
function New-Report {
    param([string]$PartialPath)
    if (Test-Path -LiteralPath $PartialPath) {
        Write-Log "Overwriting old partial report $PartialPath" 'WARN'
    }
    [IO.File]::WriteAllText($PartialPath, "=== QA Videoanalyse Protokoll ===`n`n", $Utf8Bom)
}

# AppendAllText only writes the BOM into an empty file, so it is not repeated.
function Add-ReportSection {
    param([string]$PartialPath, [int]$Number, [string]$Answer)
    [IO.File]::AppendAllText($PartialPath, "--- Abschnitt $Number ---`n$Answer`n`n", $Utf8Bom)
}

# Converts one finished report with txt2excel.ps1. A failure here is only a warning:
# the .txt report is complete, and txt2excel.bat can be run again later.
function Convert-ReportToExcel {
    param([string]$ReportPath)
    $converter = Join-Path $BaseDir 'txt2excel.ps1'
    if (-not (Test-Path -LiteralPath $converter)) {
        Write-Log "Excel conversion skipped: $converter not found" 'WARN'
        return
    }
    try {
        $global:LASTEXITCODE = 0
        & $converter -File $ReportPath
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Excel conversion failed for $ReportPath (details in logs\txt2excel.log); run txt2excel.bat later" 'WARN'
        }
    } catch {
        Write-Log "Excel conversion failed for ${ReportPath}: $($_.Exception.Message)" 'WARN'
    }
}

# Videos stay where they are. On failure writes logs\<name>.error.txt;
# on success removes an old one from an earlier failed run.
function Complete-Video {
    param($Video, [bool]$Success, [string]$Step, [string]$ErrorMessage)

    $errorFile = Join-Path $LogDir "$($Video.BaseName).error.txt"
    if ($Success) {
        if (Test-Path -LiteralPath $errorFile) { Remove-Item -LiteralPath $errorFile -Force }
        return
    }
    $lines = @(
        "Time:  $([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss', $Invariant))"
        "Video: $($Video.FullName)"
        "Step:  $Step"
        ''
        'Error:'
        $ErrorMessage
    )
    if ($script:LastToolStderr) { $lines += @('', 'Last stderr (ffmpeg/claude):', $script:LastToolStderr) }
    [IO.File]::WriteAllText($errorFile, ($lines -join "`r`n") + "`r`n", $Utf8Bom)
}

# ------------------------------------------------------------ processing

function Invoke-VideoProcessing {
    param($Video, [string]$PromptTemplate, [string]$ClaudePath)

    $name         = $Video.BaseName
    $videoWorkDir = Join-Path $WorkDir $name
    $frameDir     = Join-Path $videoWorkDir 'frames'
    $step         = 'prepare'
    $success      = $false
    $errorMessage = ''
    $reportPath   = Join-Path $ResultDir "$name.txt"
    $partialPath  = "$reportPath.partial"
    $videoCost    = 0.0
    $videoStarted = [DateTime]::Now
    $script:LastToolStderr = ''

    if (-not ($DryRun -or $TestOneWindow) -and (Test-Path -LiteralPath $reportPath)) {
        Write-Log "Skipping '$($Video.FullName)': report already exists ($reportPath)"
        $script:VideosSkipped++
        return
    }

    Write-Log "=== Video '$($Video.FullName)' ==="
    try {
        if (Test-Path -LiteralPath $videoWorkDir) {
            Write-Log "Removing leftover work folder $videoWorkDir" 'WARN'
            Remove-Item -LiteralPath $videoWorkDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $videoWorkDir -Force | Out-Null

        $step = 'ffmpeg'
        $started = [DateTime]::Now
        $frames = Invoke-FrameExtraction -VideoPath $Video.FullName -FrameDir $frameDir
        $windows = Get-Windows -FramePaths $frames
        Write-Log ([string]::Format($Invariant, 'Extracted {0} frames in {1:0.0} s, {2} windows',
            $frames.Count, ([DateTime]::Now - $started).TotalSeconds, $windows.Count))

        $step = 'task files'
        foreach ($window in $windows) {
            New-TaskFile -Window $window -VideoWorkDir $videoWorkDir -PromptTemplate $PromptTemplate | Out-Null
        }

        if ($DryRun) {
            Write-Log "DryRun: wrote $($windows.Count) task files to $videoWorkDir (no Claude calls)"
            return
        }

        if ($TestOneWindow) {
            $step = 'claude window 1'
            $window = $windows[0]
            $label = 'Window {0}/{1} ({2}-{3})' -f $window.Number, $windows.Count, $window.PrimaryStart, $window.PrimaryEnd
            Write-Log "${label}: $($window.Frames.Count) frames, calling Claude Code"
            $response = Invoke-ClaudeAnalysis -ClaudePath $ClaudePath -WorkingDir $videoWorkDir `
                -Instruction ($ClaudeInstruction -f 'auftrag_0001.txt') -Label $label -Tag '0001'
            $answer = Get-CleanAnswer -Text $response.Text -Label $label
            Write-Host ''
            Write-Host '==================== RAW ANSWER ===================='
            Write-Host $response.Text
            Write-Host '=================== CLEAN ANSWER ==================='
            Write-Host $answer
            Write-Host '===================================================='
            [IO.File]::WriteAllText((Join-Path $videoWorkDir 'antwort_0001_raw.txt'), $response.Text, $Utf8Bom)
            [IO.File]::WriteAllText((Join-Path $videoWorkDir 'antwort_0001_clean.txt'), $answer, $Utf8Bom)
            Write-Log "TestOneWindow: no report written, work folder kept: $videoWorkDir"
            return
        }

        $step = 'report'
        New-Report -PartialPath $partialPath

        foreach ($window in $windows) {
            $step = "claude window $($window.Number)"
            $label = 'Window {0}/{1} ({2}-{3})' -f $window.Number, $windows.Count, $window.PrimaryStart, $window.PrimaryEnd
            $taskName = 'auftrag_{0:0000}.txt' -f $window.Number
            Write-Log "${label}: $($window.Frames.Count) frames, calling Claude Code"
            $response = Invoke-ClaudeAnalysis -ClaudePath $ClaudePath -WorkingDir $videoWorkDir `
                -Instruction ($ClaudeInstruction -f $taskName) -Label $label -Tag ('{0:0000}' -f $window.Number)
            $answer = Get-CleanAnswer -Text $response.Text -Label $label
            Add-ReportSection -PartialPath $partialPath -Number $window.Number -Answer $answer
            if ($null -ne $response.Cost) { $videoCost += [double]$response.Cost }
        }

        $step = 'report'
        Move-Item -LiteralPath $partialPath -Destination $reportPath
        $success = $true

        if ($ConvertToExcel) {
            $step = 'excel'
            Convert-ReportToExcel -ReportPath $reportPath
        }
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Video '$($Video.Name)' failed in step '$step': $errorMessage" 'ERROR'
    } finally {
        if (-not ($DryRun -or $TestOneWindow)) {
            Complete-Video -Video $Video -Success $success -Step $step -ErrorMessage $errorMessage
            $elapsed = ([DateTime]::Now - $videoStarted).TotalMinutes
            if ($success) {
                $script:VideosDone++
                Write-Log ([string]::Format($Invariant, "Video '{0}': SUCCESS in {1:0.0} min, cost {2:0.00} USD (list price); report {3}",
                    $Video.Name, $elapsed, $videoCost, $reportPath))
            } else {
                $script:VideosFailed++
                $errorFile = Join-Path $LogDir "$name.error.txt"
                Write-Log "Video '$($Video.Name)': FAILED; details in $errorFile" 'ERROR'
            }
        }
        if (-not ($KeepWork -or $DryRun -or $TestOneWindow) -and (Test-Path -LiteralPath $videoWorkDir)) {
            try {
                Remove-Item -LiteralPath $videoWorkDir -Recurse -Force
            } catch {
                Write-Log "Could not delete work folder ${videoWorkDir}: $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Invoke-ClaudeSelfTest {
    param([string]$ClaudePath)
    $testDir = Join-Path $WorkDir '_testclaude'
    if (Test-Path -LiteralPath $testDir) { Remove-Item -LiteralPath $testDir -Recurse -Force }
    New-Item -ItemType Directory -Path $testDir -Force | Out-Null
    try {
        $r = Invoke-ClaudeAnalysis -ClaudePath $ClaudePath -WorkingDir $testDir `
            -Instruction 'Antworte nur mit: OK' -Label 'TestClaude' -Tag 'test'
        $models = @()
        $usage = Get-JsonValue $r.Json 'modelUsage'
        if ($usage) { $models = @($usage.PSObject.Properties.Name) }
        Write-Host ''
        Write-Host "result   : $($r.Text)"
        Write-Host "is_error : $(Get-JsonValue $r.Json 'is_error')"
        Write-Host "model    : $($models -join ', ')"
        Write-Host ([string]::Format($Invariant, 'duration : {0:0.0} s', $r.Seconds))
        Write-Host "cost USD : $($r.Cost)"
    } finally {
        if (-not $KeepWork) { Remove-Item -LiteralPath $testDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-VideoQa {
    $claudePath = Resolve-Claude
    if ($claudePath) {
        Write-Log "Claude Code: $claudePath"
    } elseif ($DryRun) {
        Write-Log 'Claude Code not found on PATH (not needed for -DryRun)' 'WARN'
    } else {
        throw "Claude Code ('claude') was not found on PATH for user $env:USERNAME"
    }

    if ($TestClaude) {
        Invoke-ClaudeSelfTest -ClaudePath $claudePath
        return
    }

    if (@($Video).Count -eq 0) {
        throw 'No videos given. Usage: run.bat <video file or folder> [more files or folders] [-DryRun] [-TestOneWindow] [-KeepWork]'
    }
    if (-not (Test-Path -LiteralPath $FfmpegPath)) { throw "ffmpeg not found: $FfmpegPath" }
    if (-not (Test-Path -LiteralPath $PromptFile)) { throw "Prompt file not found: $PromptFile" }
    $promptTemplate = [IO.File]::ReadAllText($PromptFile, [Text.Encoding]::UTF8)

    Write-Log "Input: $(@($Video | ForEach-Object { $_.Trim().Trim('"') }) -join ' | ')"
    $videos = Get-InputVideos -Paths $Video
    Write-Log "$($videos.Count) video(s) to process"
    if ($TestOneWindow -and $videos.Count -gt 1) {
        $videos = @($videos[0])
        Write-Log "TestOneWindow: using only '$($videos[0].FullName)'"
    }
    foreach ($item in $videos) {
        Invoke-VideoProcessing -Video $item -PromptTemplate $promptTemplate -ClaudePath $claudePath
    }
    if (-not ($DryRun -or $TestOneWindow) -and $videos.Count -gt 0) {
        Write-Log ("Summary: {0} succeeded, {1} failed, {2} skipped (report already exists); reports in {3}" -f
            $script:VideosDone, $script:VideosFailed, $script:VideosSkipped, $ResultDir)
    }
}

# ------------------------------------------------------------------ main

Initialize-Folders
$lock = Enter-RunLock
if (-not $lock) {
    Write-Log 'Another instance is already running; exiting'
    exit 0
}

$exitCode = 0
try {
    $mode = @('TestClaude', 'DryRun', 'TestOneWindow', 'KeepWork' | Where-Object { (Get-Variable $_ -ValueOnly).IsPresent }) -join ', '
    Write-Log ("Run started" + $(if ($mode) { " ($mode)" } else { '' }))
    Invoke-VideoQa
} catch {
    Write-Log "Run aborted: $($_.Exception.Message)`n$($_.ScriptStackTrace)" 'ERROR'
    $exitCode = 1
} finally {
    Write-Log 'Run finished'
    Exit-RunLock $lock
}
exit $exitCode
