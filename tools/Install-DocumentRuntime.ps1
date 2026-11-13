param([string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot))

$ErrorActionPreference = 'Stop'
$project = (Resolve-Path -LiteralPath $ProjectRoot).Path
$runtimeRoot = [IO.Path]::GetFullPath((Join-Path $project '.runtime'))
$destination = Join-Path $runtimeRoot 'libreoffice'
$version = '26.2.6.2'
$archiveName = "LibreOffice_${version}_Win_x86-64.msi"
$url = "https://downloadarchive.documentfoundation.org/libreoffice/old/$version/win/x86_64/$archiveName"
$expectedHash = '788CE7D4B56460357F57552CB8CD848E8F2254F28F6E51FA61E0C65A18096373'
$expectedBytes = 373293056
if (-not $runtimeRoot.StartsWith($project + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Runtime destination escapes the project'
}
if ((Test-Path -LiteralPath $runtimeRoot) -and ((Get-Item -LiteralPath $runtimeRoot).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'Runtime directory must not be a link'
}
if (Test-Path -LiteralPath $destination) {
    throw 'Runtime already exists; it has not been modified'
}
$downloads = Join-Path $runtimeRoot 'downloads'
New-Item -ItemType Directory -Path $downloads -Force | Out-Null
if ((Get-Item -LiteralPath $downloads).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'Download directory must not be a link'
}
$archive = Join-Path $downloads $archiveName
if (-not (Test-Path -LiteralPath $archive)) {
    $partial = $archive + '.partial'
    $offset = 0
    if (Test-Path -LiteralPath $partial) {
        $partialInfo = Get-Item -LiteralPath $partial
        if (($partialInfo.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $partialInfo.PSIsContainer) {
            throw 'Partial download must be a regular file'
        }
        $offset = $partialInfo.Length
    }
    if ($offset -gt $expectedBytes) {
        throw 'Partial download exceeds the fixed archive size; retained for inspection'
    }
    if ($offset -lt $expectedBytes) {
        Add-Type -AssemblyName System.Net.Http
        $client = [Net.Http.HttpClient]::new()
        $client.Timeout = [TimeSpan]::FromSeconds(600)
        $cancel = [Threading.CancellationTokenSource]::new(600000)
        $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Get, $url)
        $response = $null
        $inputStream = $null
        $outputStream = $null
        if ($offset -gt 0) {
            $request.Headers.Range = [Net.Http.Headers.RangeHeaderValue]::new($offset, $null)
        }
        try {
            $response = $client.SendAsync($request, [Net.Http.HttpCompletionOption]::ResponseHeadersRead, $cancel.Token).GetAwaiter().GetResult()
            if ($offset -gt 0) {
                $range = $response.Content.Headers.ContentRange
                if ([int]$response.StatusCode -ne 206 -or $null -eq $range -or $range.From -ne $offset -or $range.Length -ne $expectedBytes) {
                    throw 'Server did not confirm the requested partial range; nothing appended'
                }
            } elseif ([int]$response.StatusCode -ne 200) {
                throw 'Official archive request failed'
            }
            if ($response.Content.Headers.ContentLength -ne ($expectedBytes - $offset)) {
                throw 'Archive response length differs from the fixed release'
            }
            $inputStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
            $outputStream = [IO.FileStream]::new($partial, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::None)
            $buffer = [byte[]]::new(65536)
            $count = $offset
            while (($read = $inputStream.ReadAsync($buffer, 0, $buffer.Length, $cancel.Token).GetAwaiter().GetResult()) -gt 0) {
                $count += $read
                if ($count -gt $expectedBytes) { throw 'Archive exceeded the fixed size limit' }
                $outputStream.Write($buffer, 0, $read)
            }
            if ($count -ne $expectedBytes) { throw 'Download incomplete; bounded partial file retained' }
        } finally {
            if ($null -ne $outputStream) { $outputStream.Dispose() }
            if ($null -ne $inputStream) { $inputStream.Dispose() }
            if ($null -ne $response) { $response.Dispose() }
            $request.Dispose()
            $cancel.Dispose()
            $client.Dispose()
        }
    }
    if ((Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash -ne $expectedHash) {
        throw 'Completed download checksum mismatch; nothing extracted'
    }
    Move-Item -LiteralPath $partial -Destination $archive
}
if ((Get-Item -LiteralPath $archive).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'Archive must not be a link'
}
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $expectedHash) {
    throw 'Archive checksum mismatch; nothing extracted'
}
$arguments = '/a "' + $archive + '" /qn TARGETDIR="' + $destination + '"'
$owned = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList $arguments -WindowStyle Hidden -PassThru
try {
    if (-not $owned.WaitForExit(180000)) {
        $owned.Kill()
        $owned.WaitForExit()
        throw 'Extraction timeout; incomplete files retained for inspection'
    }
    if ($owned.ExitCode -ne 0) {
        throw "Extraction failed: $($owned.ExitCode)"
    }
} finally {
    $owned.Dispose()
}
$binary = Join-Path $destination 'program/soffice.com'
if (-not (Test-Path -LiteralPath $binary -PathType Leaf)) {
    throw 'Expected runtime executable missing'
}
$manifest = [ordered]@{
    archive_url = $url
    archive_sha256 = $expectedHash.ToLowerInvariant()
    executable_sha256 = (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash.ToLowerInvariant()
    file_version = (Get-Item -LiteralPath $binary).VersionInfo.FileVersion
    deployment = 'project-local MSI administrative extraction'
    rendering_verified = $false
}
$manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runtimeRoot 'libreoffice-manifest.json') -Encoding utf8
Write-Output "Extracted $binary; run examples/eval_fidelity.py to verify rendering."
