param([switch]$SkipBuild)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

if (-not $SkipBuild) {
    & (Join-Path $root 'build.cmd')
    if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE." }
}

$dist = Join-Path $root 'dist'
$exe = Join-Path $dist 'MeetingCapture.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw 'Build the application first.' }
$versionInfo = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
$version = '{0}.{1}.{2}' -f $versionInfo.FileMajorPart, $versionInfo.FileMinorPart, $versionInfo.FileBuildPart
$zip = Join-Path $dist "MeetingCapture-$version-windows-x64.zip"
$staging = Join-Path $dist ('.package-' + [Guid]::NewGuid().ToString('N'))

try {
    $null = New-Item -ItemType Directory -Path $staging
    Copy-Item -LiteralPath $exe -Destination (Join-Path $staging 'MeetingCapture.exe')
    Copy-Item -LiteralPath (Join-Path $root 'docs\QUICKSTART.txt') -Destination (Join-Path $staging 'README.txt')
    Copy-Item -LiteralPath (Join-Path $root 'LICENSE') -Destination $staging
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip }
    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip -CompressionLevel Optimal
} finally {
    if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
}

$lines = foreach ($file in @($exe, $zip)) {
    $hash = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant()
    '{0}  {1}' -f $hash, [IO.Path]::GetFileName($file)
}
[IO.File]::WriteAllLines((Join-Path $dist 'SHA256SUMS.txt'), $lines, [Text.Encoding]::ASCII)
Write-Output "Packaged Meeting Capture $version for Windows x64."
