param(
    [string]$Version = "0.1.0-preview.1"
)

$ErrorActionPreference = "Stop"
if ($Version -notmatch '^\d+\.\d+\.\d+(?:-[a-zA-Z0-9.-]+)?$') {
    throw "Version must be a semantic version without a leading v."
}

$root = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path -LiteralPath $root -PathType Container)) {
    throw "The project root does not exist."
}

$artifacts = Join-Path $root "artifacts"
$name = "ReplyBrief-$Version-win-x64"
$package = Join-Path $artifacts $name
$archive = Join-Path $artifacts "$name.zip"
if ((Test-Path -LiteralPath $package) -or (Test-Path -LiteralPath $archive)) {
    throw "This version is already packaged. Use a new version or explicitly remove its artifacts first."
}

New-Item -ItemType Directory -Path $package -Force | Out-Null
& dotnet publish (Join-Path $root "src\ReplyBrief.Desktop\ReplyBrief.Desktop.csproj") `
    -c Release -r win-x64 --self-contained true `
    -p:Version=$Version -p:RuntimeFrameworkVersion=8.0.31 -p:PublishTrimmed=false -p:PublishSingleFile=false `
    -p:DebugType=None -p:DebugSymbols=false -o $package
if ($LASTEXITCODE -ne 0) {
    throw "Windows publish failed."
}

if (-not (Test-Path -LiteralPath (Join-Path $package "ReplyBrief.Desktop.exe"))) {
    throw "The published application executable is missing."
}
if (-not (Test-Path -LiteralPath (Join-Path $package "outlook\manifest.xml"))) {
    throw "The published Outlook add-in is missing."
}

Copy-Item -LiteralPath (Join-Path $root "README.md") -Destination $package
Copy-Item -LiteralPath (Join-Path $root "docs") -Destination (Join-Path $package "docs") -Recurse
Compress-Archive -LiteralPath $package -DestinationPath $archive -CompressionLevel Optimal
$hash = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
[System.IO.File]::WriteAllText("$archive.sha256", "$hash  $name.zip`n", [System.Text.Encoding]::ASCII)
"Published $archive"
"SHA256: $hash"
