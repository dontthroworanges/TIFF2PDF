<#
.SYNOPSIS
    Builds "TIFF2PDF.exe" and copies it next to Convert-TiffToSearchablePdf.ps1.

.DESCRIPTION
    Produces one self-contained exe (.NET and the Windows App SDK are bundled, so users don't need to
    install anything). Requires the .NET 8 SDK or later to build. Build output goes to
    %LOCALAPPDATA%\TiffToPdfBuild so it isn't synced by OneDrive.
#>
$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot 'TiffToPdf\TiffToPdf.csproj'
$publish = Join-Path $env:LOCALAPPDATA 'TiffToPdfBuild\publish'
$target  = Join-Path (Split-Path $PSScriptRoot) 'TIFF2PDF.exe'

if (Test-Path -LiteralPath $publish) { Remove-Item -LiteralPath $publish -Recurse -Force }

dotnet publish $project -c Release -r win-x64 -o $publish `
    -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true -p:EnableCompressionInSingleFile=true
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed (exit code $LASTEXITCODE)" }

# Keep the name the project gives the exe - WinUI finds its embedded resources by the exe name.
Copy-Item -LiteralPath (Join-Path $publish (Split-Path $target -Leaf)) -Destination $target -Force
Write-Host "Published: $target ($([math]::Round((Get-Item -LiteralPath $target).Length / 1MB, 1)) MB)" -ForegroundColor Green

# The single-file exe unpacks ~190 MB to %TEMP%\.net\<exe name>\<build id> on its first launch; every build
# gets a new id. Remove earlier builds' copies (one still in use by a running copy of the app is skipped).
$extracted = Join-Path $env:TEMP ".net\$([IO.Path]::GetFileNameWithoutExtension($target))"
if (Test-Path -LiteralPath $extracted) {
    Get-ChildItem -LiteralPath $extracted -Directory | ForEach-Object {
        try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
    }
}
