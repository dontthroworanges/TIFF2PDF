<#
.SYNOPSIS
    Recursively converts .tif/.tiff files to individual searchable (OCR'd) PDFs using portable tools.

.DESCRIPTION
    Uses two portable tools kept in a tools folder (nothing is installed):
      - ImageMagick (portable build)  - prepares each page at the right resolution/compression
      - Tesseract OCR (UB-Mannheim)   - builds the PDF with an invisible, searchable text layer
    If either tool is missing from the tools folder (and not on PATH), the script offers to download
    it from the official GitHub releases. 7-Zip (portable) is also downloaded to unpack them.

    Compression is chosen per file, tuned for scanned technical drawings (line work + text):
      - Bilevel (black/white) scans  -> CCITT Group 4 (lossless, very small). Never downsampled.
      - Palette / indexed-color      -> Flate (lossless).
      - Grayscale / full color       -> JPEG, quality 90, 4:4:4 chroma (keeps line edges crisp).
                                        Use -LosslessColor for Flate instead.
    Tesseract embeds these page images as-is (no re-compression).
    Grayscale/color images above -MaxDpi are downsampled to -MaxDpi (default 400). Images with no
    resolution information are assumed to be -AssumedDpi (default 300) so the PDF page size is correct.

    Each page's orientation is checked by OCRing a reduced copy at 0/90/180/270 degrees; pages that
    read clearly better rotated are turned upright before the PDF is built (lossless for the page
    image - rotation happens before the final encode). Use -NoAutoRotate to skip this.

    OCR uses page segmentation mode 11 (sparse text) by default, which finds scattered labels,
    dimensions and title-block text on drawings better than the paragraph-oriented default.

    -Threads sets how many CPU threads are used in total. Files are converted in parallel, one file per
    thread; when there are fewer files than threads, the spare threads are shared out between the files
    and used to check orientation, rotate and compress several pages (and orientations) at once. When
    more than one file runs at a time, each file's output is shown when that file finishes.

    Each PDF is written next to its source TIFF with the same base name. Existing PDFs are skipped
    unless -Overwrite is used. A CSV log is written to the input folder unless -NoLog is used.

    "TIFF2PDF.exe" (source in the Source folder) is a Windows app that runs this script;
    it must be kept in the same folder as this script.

.EXAMPLE
    .\Convert-TiffToSearchablePdf.ps1
    Prompts for the tools folder and the folder to process.

.EXAMPLE
    .\Convert-TiffToSearchablePdf.ps1 -InputPath "D:\Scans\Drawings" -ToolsPath "C:\Tools\TiffToPdf" -Overwrite -Threads 0
    Converts using every CPU core.

.EXAMPLE
    .\Convert-TiffToSearchablePdf.ps1 -InputPath "D:\Scans\Drawings\Sheet 12.tif" -ToolsPath "C:\Tools\TiffToPdf"
    Converts a single file.
#>
[CmdletBinding()]
param(
    # Folder to search (recursively), or a single .tif/.tiff file.
    [string]$InputPath,

    # Folder holding (or to receive) the portable tools. Prompted for if not supplied.
    [string]$ToolsPath,

    # Download missing tools without asking.
    [switch]$DownloadMissingTools,

    [ValidateRange(150, 1200)]
    [int]$MaxDpi = 400,

    [ValidateRange(72, 1200)]
    [int]$AssumedDpi = 300,

    [ValidateRange(50, 100)]
    [int]$JpegQuality = 90,

    [switch]$LosslessColor,

    # Tesseract language(s), e.g. 'eng' or 'eng+spa'
    [string]$Language = 'eng',

    # 11 = sparse text (best for drawings), 3 = automatic (best for text documents), 12 = sparse + OSD
    [ValidateSet(1, 3, 4, 6, 11, 12)]
    [int]$PageSegMode = 11,

    # Skip the automatic orientation check (faster, but sideways/upside-down pages stay as scanned).
    [switch]$NoAutoRotate,

    [switch]$Overwrite,

    # CPU threads to use in total (spare threads work on pages within a file). 0 = one per CPU core.
    [ValidateRange(0, 256)]
    [int]$Threads = 1,

    # Don't write the CSV log.
    [switch]$NoLog,

    # Also print "##STATUS|<file>|<step>" lines as each file moves through the steps (used by the app
    # to show what's happening while files are still in progress).
    [switch]$ReportStatus
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest is very slow with the progress bar on

# When run by the GUI (output redirected), write UTF-8 so non-English file names come through intact.
if ([Console]::IsOutputRedirected) { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) }

$DefaultToolsPath = Join-Path $env:LOCALAPPDATA 'TiffToPdfTools'
$RotationProbeDpi = 200   # resolution used for the orientation check (lower = faster)

#region Helpers

function Invoke-Tool {
    # Runs a native command, returns exit code and combined output without tripping on stderr.
    param([string]$FilePath, [string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $output = & $FilePath @Arguments 2>&1 | ForEach-Object { "$_" }
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($output -join [Environment]::NewLine) }
}

function New-ToolStep {
    # One tool command for Invoke-ToolBatch.
    param([string]$FilePath, [string[]]$Arguments)
    [pscustomobject]@{ FilePath = $FilePath; Arguments = $Arguments }
}

function Join-ToolArgs {
    # Builds a Windows command line from separate arguments, quoted the way the C runtime parses them.
    param([string[]]$Arguments)
    $parts = foreach ($a in $Arguments) {
        if ($a.Length -gt 0 -and $a -notmatch '[\s"]') { $a; continue }
        '"' + ($a -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"'
    }
    $parts -join ' '
}

function Start-ToolProcess {
    # Starts a tool command, limiting ImageMagick/Tesseract to $ToolThreads threads of their own.
    param($Step, [int]$ToolThreads)
    $psi = New-Object System.Diagnostics.ProcessStartInfo $Step.FilePath
    $psi.Arguments = Join-ToolArgs $Step.Arguments
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = $utf8NoBom
    $psi.EnvironmentVariables['MAGICK_THREAD_LIMIT'] = "$ToolThreads"
    $psi.EnvironmentVariables['OMP_THREAD_LIMIT'] = "$ToolThreads"
    $p = [System.Diagnostics.Process]::Start($psi)
    # Read both streams in the background so a full pipe can't stall the tool.
    [pscustomobject]@{ Process = $p; StdOut = $p.StandardOutput.ReadToEndAsync(); StdErr = $p.StandardError.ReadToEndAsync() }
}

function Invoke-ToolBatch {
    # Runs independent tasks using up to $Threads threads: up to $Threads tasks at once, and each tool
    # may use its share of the threads (all of them when there's only one task). A task is one step
    # (New-ToolStep) or an array of steps run in order, stopping at the first failure. Returns one result
    # per task, in order: the exit code, stdout and combined output of its last step.
    param([object[]]$Tasks, [int]$Threads, [scriptblock]$OnProgress)
    $limit = [math]::Max(1, [math]::Min($Threads, $Tasks.Count))
    $toolThreads = [math]::Max(1, [math]::Floor($Threads / $limit))
    $results = New-Object object[] $Tasks.Count
    $running = New-Object System.Collections.Generic.List[object]
    $next = 0
    $finishedCount = 0
    try {
        while ($finishedCount -lt $Tasks.Count) {
            while ($running.Count -lt $limit -and $next -lt $Tasks.Count) {
                $steps = @($Tasks[$next])
                $running.Add([pscustomobject]@{ Index = $next; Steps = $steps; Step = 0; Tool = (Start-ToolProcess $steps[0] $toolThreads) })
                $next++
            }
            $exited = @($running | Where-Object { $_.Tool.Process.HasExited })
            if ($exited.Count -eq 0) { Start-Sleep -Milliseconds 20; continue }
            foreach ($t in $exited) {
                $p = $t.Tool.Process
                $p.WaitForExit()
                $stdout = $t.Tool.StdOut.Result
                $stderr = $t.Tool.StdErr.Result
                $code = $p.ExitCode
                $p.Dispose()
                if ($code -eq 0 -and $t.Step + 1 -lt $t.Steps.Count) {
                    $t.Step++
                    $t.Tool = Start-ToolProcess $t.Steps[$t.Step] $toolThreads
                    continue
                }
                $results[$t.Index] = [pscustomobject]@{
                    ExitCode = $code; StdOut = $stdout
                    Output   = ((@($stdout, $stderr) | Where-Object { $_ }) -join [Environment]::NewLine).Trim()
                }
                [void]$running.Remove($t)
                $finishedCount++
                if ($OnProgress) { & $OnProgress $finishedCount $Tasks.Count }
            }
        }
    }
    finally {
        foreach ($t in $running) {
            try { $t.Tool.Process.Kill() } catch { }
            $t.Tool.Process.Dispose()
        }
    }
    return , $results
}

function Invoke-FileTool {
    # Runs one tool command using all of $Threads.
    param([string]$FilePath, [string[]]$Arguments, [int]$Threads)
    (Invoke-ToolBatch @(New-ToolStep $FilePath $Arguments) $Threads)[0]
}

function Read-YesNo {
    param([string]$Prompt)
    while ($true) {
        $a = (Read-Host "$Prompt (Y/N)").Trim()
        if ($a -match '^(y|yes)$') { return $true }
        if ($a -match '^(n|no)$')  { return $false }
    }
}

function Read-FolderPath {
    param([string]$Prompt, [string]$Default, [switch]$MustExist)
    while ($true) {
        $text = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
        $p = (Read-Host $text).Trim().Trim('"').Trim("'")
        if (-not $p) { $p = $Default }
        if (-not $p) { continue }
        if (Test-Path -LiteralPath $p -PathType Container) { return (Resolve-Path -LiteralPath $p).ProviderPath }
        if (-not $MustExist) { return [IO.Path]::GetFullPath($p) }
        Write-Host "Folder not found: '$p'. Please try again." -ForegroundColor Yellow
    }
}

function Find-ToolExe {
    # Looks for an exe in the tools folder (up to 2 levels deep), then on PATH.
    param([string]$ExeName)
    if (Test-Path -LiteralPath $ToolsPath) {
        $hit = Get-ChildItem -LiteralPath $ToolsPath -Filter $ExeName -Recurse -Depth 2 -File -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    $cmd = Get-Command $ExeName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) { return $cmd.Source }
    return $null
}

function Get-GitHubReleaseAsset {
    # Downloads the first asset of the latest release of $Repo whose name matches $Pattern.
    param([string]$Repo, [string]$Pattern, [string]$OutDir)
    $headers = @{ 'User-Agent' = 'Convert-TiffToSearchablePdf' }
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" -Headers $headers
    $asset = $release.assets | Where-Object { $_.name -match $Pattern } | Select-Object -First 1
    if (-not $asset) { throw "No download matching '$Pattern' found in the latest $Repo release." }

    $dest = Join-Path $OutDir $asset.name
    Write-Host ("    Downloading {0} ({1:N1} MB)..." -f $asset.name, ($asset.size / 1MB))
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $dest -Headers $headers -UseBasicParsing
    Unblock-File -LiteralPath $dest
    return $dest
}

function Get-SevenZip {
    # Ensures a full portable 7-Zip (7z.exe + 7z.dll) exists in the tools folder; needed to unpack
    # the ImageMagick .7z and the Tesseract installer. 7zr.exe alone can only open .7z files, so it is
    # used to unpack the full 7-Zip from its installer (which is a 7z archive).
    param([string]$DownloadDir)
    $sevenZipDir = Join-Path $ToolsPath '7-Zip'
    $sevenZip = Join-Path $sevenZipDir '7z.exe'
    if (Test-Path -LiteralPath $sevenZip) { return $sevenZip }

    Write-Host '  7-Zip (portable, used to unpack the other tools)' -ForegroundColor Cyan
    $sevenZr    = Get-GitHubReleaseAsset 'ip7z/7zip' '^7zr\.exe$'        $DownloadDir
    $sevenSetup = Get-GitHubReleaseAsset 'ip7z/7zip' '^7z\d+-x64\.exe$'  $DownloadDir
    $r = Invoke-Tool $sevenZr @('x', $sevenSetup, "-o$sevenZipDir", '-y')
    if ($r.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $sevenZip)) { throw "Could not unpack 7-Zip: $($r.Output)" }
    return $sevenZip
}

function Install-PortableTools {
    param([bool]$NeedMagick, [bool]$NeedTesseract)

    New-Item -ItemType Directory -Force -Path $ToolsPath | Out-Null
    $downloadDir = Join-Path $ToolsPath '_downloads'
    New-Item -ItemType Directory -Force -Path $downloadDir | Out-Null

    # Help with corporate proxies (PowerShell 5.1) and older TLS defaults.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    if ([Net.WebRequest]::DefaultWebProxy) {
        [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
    }

    try {
        $sevenZip = Get-SevenZip $downloadDir

        if ($NeedMagick) {
            Write-Host '  ImageMagick (portable)' -ForegroundColor Cyan
            $archive = Get-GitHubReleaseAsset 'ImageMagick/ImageMagick' '^ImageMagick-[\d\.\-]+-portable-Q16-x64\.7z$' $downloadDir
            $r = Invoke-Tool $sevenZip @('x', $archive, "-o$(Join-Path $ToolsPath 'ImageMagick')", '-y')
            if ($r.ExitCode -ne 0) { throw "Could not unpack ImageMagick: $($r.Output)" }
        }

        if ($NeedTesseract) {
            Write-Host '  Tesseract OCR (unpacked from the UB-Mannheim installer, not installed)' -ForegroundColor Cyan
            $setup = Get-GitHubReleaseAsset 'UB-Mannheim/tesseract' '^tesseract-ocr-w64-setup-.*\.exe$' $downloadDir
            $tessDir = Join-Path $ToolsPath 'Tesseract'
            $r = Invoke-Tool $sevenZip @('x', $setup, "-o$tessDir", '-y')
            if ($r.ExitCode -ne 0) { throw "Could not unpack Tesseract: $($r.Output)" }
            # Installer leftovers that aren't needed for a portable copy.
            Remove-Item -LiteralPath (Join-Path $tessDir '$PLUGINSDIR') -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath (Join-Path $tessDir 'tesseract-uninstall.exe') -Force -ErrorAction SilentlyContinue
        }
    }
    finally {
        Remove-Item -LiteralPath $downloadDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Install-TesseractLanguage {
    param([string]$Lang, [string]$TessDataDir)
    $url  = "https://github.com/tesseract-ocr/tessdata/raw/main/$Lang.traineddata"
    $dest = Join-Path $TessDataDir "$Lang.traineddata"
    Write-Host "    Downloading $Lang.traineddata..."
    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
}

function Get-OcrScore {
    # Amount of text (characters in words of 3+ letters/digits) recognised with >= 60% confidence,
    # from Tesseract's TSV output.
    param([string]$Tsv)
    $score = 0
    foreach ($line in $Tsv -split "`r?`n") {
        $cols = "$line".Split("`t")
        if ($cols.Count -lt 12) { continue }
        $conf = 0.0
        if (-not [double]::TryParse($cols[10], [Globalization.NumberStyles]::Float,
                                    [Globalization.CultureInfo]::InvariantCulture, [ref]$conf)) { continue }
        $word = $cols[11] -replace '[^\p{L}\p{N}]', ''
        if ($conf -ge 60 -and $word.Length -ge 3) { $score += $word.Length }
    }
    return $score
}

function Get-PageRotations {
    # Returns, for each page, the clockwise rotation (0/90/180/270) that makes its text read best.
    # A reduced-resolution copy is OCR'd in all four orientations and the one with the most
    # confidently-recognised text wins. Tesseract's built-in orientation detection (--psm 0) isn't
    # used because it often mistakes upright ALL-CAPS text (typical of drawings) for upside down.
    # A page is only rotated when the winner is clearly better than the page as scanned.
    # The pages, and the four orientations of each, are checked in parallel using $Threads.
    param([string[]]$PagePaths, [double]$Dpi, [string]$WorkDir, [int]$Threads, [scriptblock]$OnProgress)
    $degrees = 0, 90, 180, 270
    $scale = [math]::Min(100, [math]::Round(100 * $RotationProbeDpi / $Dpi))

    # One reduced grayscale copy per page; the other orientations are rotated from it (lossless).
    $probes = @(for ($p = 0; $p -lt $PagePaths.Count; $p++) { Join-Path $WorkDir "probe-$p.png" })
    $tasks = New-Object System.Collections.Generic.List[object]
    for ($p = 0; $p -lt $PagePaths.Count; $p++) {
        $tasks.Add((New-ToolStep $Magick @('-quiet', $PagePaths[$p], '-resize', "$scale%", '-colorspace', 'Gray', $probes[$p])))
    }
    $made = Invoke-ToolBatch $tasks.ToArray() $Threads

    $tasks.Clear()
    for ($p = 0; $p -lt $PagePaths.Count; $p++) {
        foreach ($deg in $degrees) {
            $image = if ($deg -eq 0) { $probes[$p] } else { Join-Path $WorkDir "probe-$p-$deg.png" }
            $ocr = New-ToolStep $Tesseract (@($image, 'stdout') + $tessDataArgs + @('-l', $Language, '--psm', "$PageSegMode", 'tsv'))
            if ($deg -eq 0) { $tasks.Add($ocr) }
            else { $tasks.Add(@((New-ToolStep $Magick @('-quiet', $probes[$p], '-rotate', "$deg", $image)), $ocr)) }
        }
    }
    $read = Invoke-ToolBatch $tasks.ToArray() $Threads $OnProgress
    Remove-Item -Path (Join-Path $WorkDir 'probe-*.png') -Force -ErrorAction SilentlyContinue

    $rotations = for ($p = 0; $p -lt $PagePaths.Count; $p++) {
        if ($made[$p].ExitCode -ne 0) { 0; continue }
        $scores = @{}
        for ($d = 0; $d -lt 4; $d++) {
            $r = $read[$p * 4 + $d]
            $scores[$degrees[$d]] = if ($r.ExitCode -eq 0) { Get-OcrScore $r.StdOut } else { 0 }
        }
        $best = $degrees | Sort-Object { $scores[$_] } -Descending | Select-Object -First 1
        if ($best -ne 0 -and $scores[$best] -ge 20 -and $scores[$best] -ge 2 * $scores[0]) { $best } else { 0 }
    }
    return , [int[]]@($rotations)
}

function Get-TiffInfo {
    # Reads type and resolution of the first frame.
    param([string]$Path)
    $r = Invoke-Tool $Magick @('identify', '-quiet', '-format', '%[type]|%x|%U|%[colorspace]', "$Path[0]")
    if ($r.ExitCode -ne 0) { throw "identify failed: $($r.Output)" }
    $line  = ($r.Output -split "`r?`n" | Where-Object { $_ -match '\|' } | Select-Object -First 1)
    $parts = $line.Split('|')
    $xRes  = 0.0
    [void][double]::TryParse(($parts[1] -replace '[^\d\.]', ''), [Globalization.NumberStyles]::Float,
                             [Globalization.CultureInfo]::InvariantCulture, [ref]$xRes)
    $units = $parts[2]

    # ImageMagick reports 72 / Undefined when the TIFF has no resolution tag.
    $hasDpi = -not ($xRes -le 1 -or ($units -eq 'Undefined' -and $xRes -eq 72))
    $dpi = if ($units -eq 'PixelsPerCentimeter') { $xRes * 2.54 } else { $xRes }

    [pscustomobject]@{
        Type       = $parts[0]
        Dpi        = [math]::Round($dpi)
        HasDpi     = $hasDpi
        ColorSpace = $parts[3]
    }
}

function New-ResultRecord {
    param($Job)
    [ordered]@{
        Source = $Job.Source; Output = $Job.Output; Status = ''; Type = ''; SourceDpi = ''; Pages = ''; Rotated = ''
        Compression = ''; SourceMB = [math]::Round($Job.Length / 1MB, 2); OutputMB = ''; Message = ''
    }
}

function Set-Step {
    # Reports which step a file is on (-ReportStatus). Worker runspaces queue it for the main loop to
    # print straight away; a single-threaded run prints it directly.
    param($Job, [string]$Step)
    if ($StatusQueue)      { $StatusQueue.Enqueue("$($Job.Source)|$Step") }
    elseif ($ReportStatus) { Write-Host "##STATUS|$($Job.Source)|$Step" }
}

function Convert-TiffFile {
    # Converts one TIFF to a searchable PDF and returns its log record. Runs either in this script
    # or in a worker runspace, so it only uses its parameter, the helper functions above and the
    # shared settings variables (see $sharedVars in Main).
    param($Job)
    $ErrorActionPreference = 'Stop'
    $result = New-ResultRecord $Job
    $threads = $Job.Threads   # this file's share of the CPU threads

    # Work on a temp copy: avoids OneDrive locks and ImageMagick treating [ ] in names as frame specifiers.
    $workDir = $Job.WorkDir
    New-Item -ItemType Directory -Path $workDir | Out-Null
    $tmpSrc = Join-Path $workDir ("src" + $Job.Extension)

    try {
        Set-Step $Job 'Reading file'
        Copy-Item -LiteralPath $Job.Source -Destination $tmpSrc -Force

        # ---- Step 1: prepare page images ----
        $info = Get-TiffInfo $tmpSrc
        $result.Type = $info.Type
        $result.SourceDpi = if ($info.HasDpi) { $info.Dpi } else { "none (assumed $AssumedDpi)" }
        $dpi = if ($info.HasDpi) { $info.Dpi } else { $AssumedDpi }

        $magickArgs = @('-quiet', $tmpSrc, '-units', 'PixelsPerInch')
        if (-not $info.HasDpi) { $magickArgs += @('-density', "$AssumedDpi") }

        # Pages are first written as lossless TIFFs so any rotation happens before the single
        # (possibly lossy) final encode.
        $isBilevel    = $info.Type -eq 'Bilevel'
        $useJpeg      = -not $isBilevel -and -not ($info.Type -like 'Palette*' -or $LosslessColor)
        $tiffCompress = if ($isBilevel) { 'Group4' } else { 'Zip' }
        $outDpi       = $dpi

        if ($isBilevel) {
            # Lossless CCITT G4 - ideal for B/W line drawings; no downsampling.
            $result.Compression = 'CCITT G4'
        } else {
            if ($dpi -gt $MaxDpi) {
                $magickArgs += @('-resample', "$MaxDpi")
                $outDpi = $MaxDpi
            }
            $magickArgs += @('-background', 'white', '-alpha', 'remove', '-alpha', 'off')
            if ($info.ColorSpace -eq 'CMYK') { $magickArgs += @('-colorspace', 'sRGB') }

            $result.Compression = if ($useJpeg) { "JPEG q$JpegQuality" } else { 'Flate (lossless)' }
            if ($dpi -gt $MaxDpi) { $result.Compression += " @ $MaxDpi dpi" }
        }
        $magickArgs += @('-compress', $tiffCompress, (Join-Path $workDir 'page-%04d.tif'))

        Write-Host "    Preparing pages ($($info.Type), $($result.Compression))..."
        Set-Step $Job 'Preparing pages'
        $r = Invoke-FileTool $Magick $magickArgs $threads
        $pages = @(Get-ChildItem -LiteralPath $workDir -Filter 'page-*.tif' | Sort-Object Name)
        if ($r.ExitCode -ne 0 -or $pages.Count -eq 0) { throw "ImageMagick failed: $($r.Output)" }
        $result.Pages = $pages.Count
        $pagePaths = [string[]]@($pages.FullName)

        # ---- Step 2: detect and fix page orientation (pages in parallel) ----
        if (-not $NoAutoRotate) {
            Write-Host '    Checking page orientation...'
            Set-Step $Job 'Checking page orientation'
            $onProgress = $null
            if ($pagePaths.Count -gt 1) {
                $pagesChecked = @{ Count = 0 }
                $onProgress = {
                    param($finished, $total)
                    $checked = [math]::Floor($finished / 4)   # four orientations per page
                    if ($checked -ne $pagesChecked.Count) {
                        $pagesChecked.Count = $checked
                        Set-Step $Job "Checking page orientation ($checked of $($pagePaths.Count) pages done)"
                    }
                }
            }
            $degrees = Get-PageRotations $pagePaths $outDpi $workDir $threads $onProgress

            $tasks = New-Object System.Collections.Generic.List[object]
            $rotations = @()
            for ($p = 0; $p -lt $pagePaths.Count; $p++) {
                if ($degrees[$p] -eq 0) { continue }
                $tasks.Add((New-ToolStep $Magick @('-quiet', $pagePaths[$p], '-rotate', "$($degrees[$p])", '-compress', $tiffCompress, $pagePaths[$p])))
                $rotations += "p$($p + 1): $($degrees[$p])"
            }
            if ($tasks.Count -gt 0) {
                $failed = Invoke-ToolBatch $tasks.ToArray() $threads | Where-Object { $_.ExitCode -ne 0 } | Select-Object -First 1
                if ($failed) { throw "Rotation failed: $($failed.Output)" }
                $result.Rotated = $rotations -join '; '
                Write-Host "    Rotated clockwise - $($result.Rotated)" -ForegroundColor Yellow
            }
        }

        # ---- Step 3: final encode for grayscale/color pages (pages in parallel) ----
        if ($useJpeg) {
            Set-Step $Job 'Compressing pages'
            $tasks = New-Object System.Collections.Generic.List[object]
            foreach ($page in $pagePaths) {
                $jpg = [IO.Path]::ChangeExtension($page, '.jpg')
                $tasks.Add((New-ToolStep $Magick @('-quiet', $page, '-quality', "$JpegQuality", '-sampling-factor', '4:4:4', $jpg)))
            }
            $failed = Invoke-ToolBatch $tasks.ToArray() $threads | Where-Object { $_.ExitCode -ne 0 } | Select-Object -First 1
            if ($failed) { throw "JPEG encode failed: $($failed.Output)" }
            $pages = @(Get-ChildItem -LiteralPath $workDir -Filter 'page-*.jpg' | Sort-Object Name)
        }

        # ---- Step 4: build searchable PDF with Tesseract ----
        Write-Host "    Running OCR on $($pages.Count) page(s)..."
        Set-Step $Job 'Running OCR'
        $listFile = Join-Path $workDir 'pages.txt'
        [IO.File]::WriteAllLines($listFile, [string[]]$pages.FullName, $utf8NoBom)
        $outBase = Join-Path $workDir 'out'
        $outPdf  = "$outBase.pdf"

        $tessArgs = @($listFile, $outBase) + $tessDataArgs + @('-l', $Language, '--psm', "$PageSegMode", 'pdf')
        $r = Invoke-FileTool $Tesseract $tessArgs $threads

        if ($r.ExitCode -eq 0 -and (Test-Path -LiteralPath $outPdf)) {
            Move-Item -LiteralPath $outPdf -Destination $Job.Output -Force
            $result.Status = 'Success'
            Write-Host '    Done.' -ForegroundColor Green
        } else {
            # OCR failed - still produce an image-only PDF so the conversion isn't lost.
            $fallbackArgs = @('-quiet') + [string[]]$pages.FullName
            if ($useJpeg) {
                $fallbackArgs += @('-compress', 'JPEG', '-quality', "$JpegQuality", '-sampling-factor', '4:4:4')
            } else {
                $fallbackArgs += @('-compress', $tiffCompress)
            }
            $fallbackArgs += $outPdf
            Set-Step $Job 'Building image-only PDF'
            $f = Invoke-FileTool $Magick $fallbackArgs $threads
            if ($f.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $outPdf)) {
                throw "Tesseract failed ($($r.ExitCode)): $($r.Output)"
            }
            Move-Item -LiteralPath $outPdf -Destination $Job.Output -Force
            $result.Status = 'Converted - OCR failed'
            $result.Message = "Tesseract exit code $($r.ExitCode): $($r.Output)"
            Write-Host "    PDF created but OCR failed (exit code $($r.ExitCode))." -ForegroundColor Yellow
        }
        $result.OutputMB = [math]::Round((Get-Item -LiteralPath $Job.Output).Length / 1MB, 2)
    }
    catch {
        $result.Status = 'Failed'
        $result.Message = $_.Exception.Message
        Write-Host "    FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    [pscustomobject]$result
}

function Write-HostRecord {
    # Replays a Write-Host message captured from a worker runspace, keeping its colour.
    param($Record)
    $m = $Record.MessageData
    if ($m -is [System.Management.Automation.HostInformationMessage]) {
        $p = @{ Object = $m.Message; NoNewline = $m.NoNewLine }
        if ($null -ne $m.ForegroundColor) { $p.ForegroundColor = $m.ForegroundColor }
        Write-Host @p
    } else {
        Write-Host "$m"
    }
}

#endregion

#region Tools

if (-not $ToolsPath) {
    Write-Host 'Portable tools folder (ImageMagick + Tesseract). Press Enter to accept the default.'
    $ToolsPath = Read-FolderPath -Prompt 'Tools folder' -Default $DefaultToolsPath
}
$ToolsPath = [IO.Path]::GetFullPath($ToolsPath)

foreach ($od in @($env:OneDrive, $env:OneDriveCommercial) | Where-Object { $_ }) {
    if ($ToolsPath.StartsWith($od, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Host 'Warning: the tools folder is inside OneDrive/SharePoint. The tools (~500 MB) will be synced.' -ForegroundColor Yellow
        break
    }
}

$Magick    = Find-ToolExe 'magick.exe'
$Tesseract = Find-ToolExe 'tesseract.exe'

if (-not $Magick -or -not $Tesseract) {
    $missing = @()
    if (-not $Magick)    { $missing += 'ImageMagick' }
    if (-not $Tesseract) { $missing += 'Tesseract OCR' }
    Write-Host "`nMissing tool(s): $($missing -join ', ')" -ForegroundColor Yellow
    Write-Host "They can be downloaded (portable, no install or admin rights needed) to:`n    $ToolsPath"

    if (-not ($DownloadMissingTools -or (Read-YesNo 'Download them now?'))) {
        Write-Host 'Cannot continue without the required tools.' -ForegroundColor Red
        exit 1
    }
    Install-PortableTools -NeedMagick (-not $Magick) -NeedTesseract (-not $Tesseract)

    $Magick    = Find-ToolExe 'magick.exe'
    $Tesseract = Find-ToolExe 'tesseract.exe'
    if (-not $Magick -or -not $Tesseract) { throw 'Tools still not found after download.' }
    Write-Host "Tools ready.`n" -ForegroundColor Green
}

# Point Tesseract at the tessdata folder next to it (portable copies don't have it registered).
$tessDataDir = Join-Path (Split-Path $Tesseract) 'tessdata'
$tessDataArgs = @()
if (Test-Path -LiteralPath $tessDataDir) { $tessDataArgs = @('--tessdata-dir', $tessDataDir) }

Write-Host "ImageMagick: $Magick"
Write-Host "Tesseract:   $Tesseract"

# Confirm the requested OCR language pack(s) are present; offer to download any that aren't.
$installedLangs = (Invoke-Tool $Tesseract ($tessDataArgs + @('--list-langs'))).Output -split "`r?`n" | ForEach-Object { $_.Trim() }
$missingLangs = @($Language.Split('+') | Where-Object { $installedLangs -notcontains $_ })
if ($missingLangs.Count -gt 0) {
    Write-Host "Tesseract language data missing: $($missingLangs -join ', ')" -ForegroundColor Yellow
    if ((Test-Path -LiteralPath $tessDataDir) -and ($DownloadMissingTools -or (Read-YesNo 'Download it now?'))) {
        foreach ($lang in $missingLangs) { Install-TesseractLanguage $lang $tessDataDir }
    } else {
        throw "Language data not available: $($missingLangs -join ', ')"
    }
}

#endregion

#region Main

if (-not $InputPath) {
    $InputPath = Read-FolderPath -Prompt "`nEnter the folder path to search for TIFF files" -MustExist
} elseif (-not (Test-Path -LiteralPath $InputPath)) {
    throw "Folder or file not found: $InputPath"
} else {
    $InputPath = (Resolve-Path -LiteralPath $InputPath).ProviderPath
}

if (Test-Path -LiteralPath $InputPath -PathType Leaf) {
    # A single file.
    $tiffs = @(Get-Item -LiteralPath $InputPath)
    if ($tiffs[0].Extension -notin '.tif', '.tiff') { throw "Not a TIFF file: $InputPath" }
    $logDir = $tiffs[0].DirectoryName
    Write-Host "`nProcessing single file '$InputPath'" -ForegroundColor Cyan
} else {
    $logDir = $InputPath
    Write-Host "`nSearching '$InputPath' for TIFF files..." -ForegroundColor Cyan
    $tiffs = @(Get-ChildItem -LiteralPath $InputPath -Recurse -File -ErrorAction SilentlyContinue |
               Where-Object { $_.Extension -in '.tif', '.tiff' } |
               Sort-Object FullName)
}

if ($tiffs.Count -eq 0) {
    Write-Host 'No .tif or .tiff files found.' -ForegroundColor Yellow
    exit 0
}
Write-Host "Found $($tiffs.Count) TIFF file(s).`n"

$workRoot = Join-Path ([IO.Path]::GetTempPath()) ("tiff2pdf_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workRoot | Out-Null

$results       = New-Object System.Collections.Generic.List[object]
$jobs          = New-Object System.Collections.Generic.List[object]
$claimedOutput = @{}
$utf8NoBom     = New-Object System.Text.UTF8Encoding($false)
$i = 0

# Decide each file's output name up front (so parallel workers can't collide) and skip existing PDFs.
foreach ($tiff in $tiffs) {
    $i++
    $baseOut = Join-Path $tiff.DirectoryName ($tiff.BaseName + '.pdf')

    # Avoid collisions when both "X.tif" and "X.tiff" exist in the same folder.
    if ($claimedOutput.ContainsKey($baseOut.ToLower())) {
        $baseOut = Join-Path $tiff.DirectoryName ("{0}_{1}.pdf" -f $tiff.BaseName, $tiff.Extension.TrimStart('.'))
    }
    $claimedOutput[$baseOut.ToLower()] = $true

    $job = [pscustomobject]@{
        Source = $tiff.FullName; Extension = $tiff.Extension; Length = $tiff.Length
        Output = $baseOut; WorkDir = (Join-Path $workRoot $i); Threads = 1
    }

    if ((Test-Path -LiteralPath $baseOut) -and -not $Overwrite) {
        $result = New-ResultRecord $job
        $result.Status = 'Skipped'
        $result.Message = 'PDF already exists (use -Overwrite to replace)'
        Write-Host "Skipped - PDF already exists: $($tiff.FullName)" -ForegroundColor DarkGray
        $results.Add([pscustomobject]$result)
        continue
    }
    $jobs.Add($job)
}

# One file per thread at a time. With fewer files than threads, the spare threads are shared out
# between the files (e.g. 8 threads, 3 files: 3 + 3 + 2) and used for the pages within each file.
$threadCount = if ($Threads -eq 0) { [Environment]::ProcessorCount } else { $Threads }
$workers     = [math]::Max(1, [math]::Min($threadCount, $jobs.Count))
$perFile     = [math]::Floor($threadCount / $workers)
$extra       = $threadCount % $workers
for ($j = 0; $j -lt $jobs.Count; $j++) {
    $jobs[$j].Threads = $perFile + $(if (($j % $workers) -lt $extra) { 1 } else { 0 })
}
if ($jobs.Count -gt 0) {
    Write-Host ("`nConverting {0} file(s) using {1} thread(s) ({2} file(s) at a time)...`n" -f $jobs.Count, $threadCount, $workers) -ForegroundColor Cyan
}

$pool        = $null
$running     = New-Object System.Collections.Generic.List[object]
$StatusQueue = $null

try {
    if ($workers -eq 1) {
        $n = 0
        foreach ($job in $jobs) {
            $n++
            Write-Host ("[{0}/{1}] {2}" -f $n, $jobs.Count, $job.Source)
            $results.Add((Convert-TiffFile $job))
        }
    } else {
        # Worker runspaces get copies of the helper functions and the settings they use.
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        if ($ReportStatus) { $StatusQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' }
        foreach ($fn in 'Invoke-Tool', 'New-ToolStep', 'Join-ToolArgs', 'Start-ToolProcess', 'Invoke-ToolBatch', 'Invoke-FileTool',
                        'Get-OcrScore', 'Get-PageRotations', 'Get-TiffInfo', 'New-ResultRecord', 'Set-Step', 'Convert-TiffFile') {
            $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry $fn, (Get-Item "function:$fn").Definition))
        }
        $sharedVars = @{
            Magick = $Magick; Tesseract = $Tesseract; tessDataArgs = $tessDataArgs; Language = $Language
            PageSegMode = $PageSegMode; AssumedDpi = $AssumedDpi; MaxDpi = $MaxDpi; JpegQuality = $JpegQuality
            LosslessColor = [bool]$LosslessColor; NoAutoRotate = [bool]$NoAutoRotate
            RotationProbeDpi = $RotationProbeDpi; utf8NoBom = $utf8NoBom; StatusQueue = $StatusQueue
        }
        foreach ($name in $sharedVars.Keys) {
            $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry $name, $sharedVars[$name], $null))
        }

        # No host is given, so Write-Host output in the workers is buffered in their Information stream.
        $pool = [runspacefactory]::CreateRunspacePool($iss)
        [void]$pool.SetMaxRunspaces($workers)
        $pool.Open()
        foreach ($job in $jobs) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddCommand('Convert-TiffFile').AddParameter('Job', $job)
            $running.Add([pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Job = $job })
        }

        # Report each file as it finishes, replaying its buffered output as one block.
        $n = 0
        $status = $null
        while ($running.Count -gt 0) {
            # Step updates first, so a file's last update is printed before its finished block.
            if ($StatusQueue) { while ($StatusQueue.TryDequeue([ref]$status)) { Write-Host "##STATUS|$status" } }
            $finished = @($running | Where-Object { $_.Handle.IsCompleted })
            if ($finished.Count -eq 0) { Start-Sleep -Milliseconds 250; continue }
            foreach ($w in $finished) {
                $n++
                Write-Host ("[{0}/{1}] {2}" -f $n, $jobs.Count, $w.Job.Source)
                try {
                    foreach ($rec in $w.PS.Streams.Information) { Write-HostRecord $rec }
                    $out = $w.PS.EndInvoke($w.Handle) | Select-Object -Last 1
                    if (-not $out) { throw ($w.PS.Streams.Error | Select-Object -First 1) }
                    $results.Add($out)
                }
                catch {
                    $result = New-ResultRecord $w.Job
                    $result.Status = 'Failed'
                    $result.Message = "$_"
                    Write-Host "    FAILED: $_" -ForegroundColor Red
                    $results.Add([pscustomobject]$result)
                }
                finally {
                    $w.PS.Dispose()
                    [void]$running.Remove($w)
                }
            }
        }
    }
}
finally {
    foreach ($w in $running) { $w.PS.Dispose() }
    if ($pool) { $pool.Dispose() }
    Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
}

#endregion

#region Summary

Write-Host "`n===== Summary =====" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object { Write-Host ("{0,-24} {1}" -f $_.Name, $_.Count) }

if (-not $NoLog) {
    $logPath = Join-Path $logDir ("TIFF_to_PDF_log_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
    $results | Sort-Object Source | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
    Write-Host "Log: $logPath"
}

exit 0

#endregion
