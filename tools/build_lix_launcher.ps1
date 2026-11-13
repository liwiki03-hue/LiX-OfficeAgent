$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) {
    $compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
}
if (-not (Test-Path -LiteralPath $compiler)) { throw 'Windows .NET Framework compiler is unavailable' }
$versionMatch = [regex]::Match([IO.File]::ReadAllText((Join-Path $projectRoot 'pyproject.toml')), '(?m)^version = "(\d+\.\d+\.\d+)"\r?$')
if (-not $versionMatch.Success) { throw 'Project version is missing' }
$buildRoot = Join-Path ([IO.Path]::GetTempPath()) ('lix-build-' + [guid]::NewGuid().ToString('N'))
$originalTemp = $env:TEMP
$originalTmp = $env:TMP
try {
    [IO.Directory]::CreateDirectory($buildRoot) | Out-Null
    $env:TEMP = $buildRoot
    $env:TMP = $buildRoot
    $metadata = Join-Path $buildRoot 'version.cs'
    $metadataText = 'using System.Reflection;[assembly: AssemblyTitle("LiX")][assembly: AssemblyProduct("LiX")][assembly: AssemblyDescription("LiX desktop launcher")][assembly: AssemblyVersion("' + $versionMatch.Groups[1].Value + '.0")]'
    [IO.File]::WriteAllText($metadata, $metadataText, [Text.UTF8Encoding]::new($false))
    $executable = Join-Path $buildRoot 'LiX.exe'
    & $compiler /nologo /target:winexe /platform:anycpu /optimize+ /reference:System.Windows.Forms.dll "/out:$executable" "/win32icon:$projectRoot/office_agent/webui/static/brand/lix.ico" (Join-Path $projectRoot 'tools/lix_launcher.cs') $metadata
    if ($LASTEXITCODE -ne 0) { throw "C# compiler failed: $LASTEXITCODE" }
    $destination = Join-Path $projectRoot 'LiX.exe'
    $previous = Join-Path $buildRoot 'LiX.previous.exe'
    if (Test-Path -LiteralPath $destination) { [IO.File]::Move($destination, $previous) }
    try { [IO.File]::Move($executable, $destination) }
    catch {
        if (Test-Path -LiteralPath $previous) { [IO.File]::Move($previous, $destination) }
        throw
    }
    Write-Output $destination
}
finally {
    $env:TEMP = $originalTemp
    $env:TMP = $originalTmp
    if ([IO.Path]::GetFullPath($buildRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($buildRoot).StartsWith('lix-build-')) {
        Remove-Item -LiteralPath $buildRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
