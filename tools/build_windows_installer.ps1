param(
    [Parameter(Mandatory=$true)][string]$Payload,
    [Parameter(Mandatory=$true)][string]$Output,
    [Parameter(Mandatory=$true)][string]$UvArchive,
    [Parameter(Mandatory=$true)][string]$LicenseRoot,
    [Parameter(Mandatory=$true)][string]$InnoSetup
)
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$originalTemp = $env:TEMP
$originalTmp = $env:TMP
$temporaryParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$buildRoot = Join-Path $temporaryParent ('lix-installer-build-' + [guid]::NewGuid().ToString('N'))
$utf8 = [Text.UTF8Encoding]::new($false)

function Get-CheckedPath([string]$Value) {
    $resolved = [IO.Path]::GetFullPath($Value)
    if ($resolved.StartsWith('\\')) { throw 'Network build paths are not supported' }
    $cursor = $resolved
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            if (([IO.File]::GetAttributes($cursor) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'Build path contains a link or reparse point'
            }
        }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    return $resolved
}

function Invoke-OwnedBuild([string]$Executable, [string[]]$Arguments, [int]$Seconds) {
    $childEnvironment = [Collections.Generic.SortedDictionary[string,string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in [Environment]::GetEnvironmentVariables().GetEnumerator()) {
        $childEnvironment[[string]$row.Key] = [string]$row.Value
    }
    $exitCode = [OwnedCommand]::Run($Executable, $Arguments, $projectRoot, $childEnvironment, $Seconds,
        [Func[bool]]{ return $false }, [Action[bool]]{ param($empty) if (-not $empty) { throw 'Build process cleanup was not confirmed' } })
    if ($exitCode -ne 0) { throw ('Build command failed with exit code ' + $exitCode + ': ' + [IO.Path]::GetFileName($Executable)) }
}

$Payload = Get-CheckedPath $Payload
$Output = Get-CheckedPath $Output
$UvArchive = Get-CheckedPath $UvArchive
$LicenseRoot = Get-CheckedPath $LicenseRoot
$InnoSetup = Get-CheckedPath $InnoSetup
Get-CheckedPath $temporaryParent | Out-Null
if ((Test-Path -LiteralPath $Payload) -or (Test-Path -LiteralPath $Output)) { throw 'Build requires new payload and output directories' }
if ([IO.Path]::GetPathRoot($buildRoot) -ne [IO.Path]::GetPathRoot($Output) -or
    [IO.Path]::GetPathRoot($buildRoot) -ne [IO.Path]::GetPathRoot($Payload)) {
    throw 'Build and output must share a volume so final binaries can be moved without copying metadata'
}
$spec = [IO.File]::ReadAllText((Join-Path $projectRoot 'release/WINDOWS_INSTALLER.json')) | ConvertFrom-Json
if ((Get-Item -LiteralPath $InnoSetup).Length -ne $spec.inno.bytes -or
    (Get-FileHash -LiteralPath $InnoSetup -Algorithm SHA256).Hash.ToLowerInvariant() -ne $spec.inno.sha256) {
    throw 'Inno Setup archive differs from the pinned release'
}
try {
    [IO.Directory]::CreateDirectory($buildRoot) | Out-Null
    $env:TEMP = $buildRoot
    $env:TMP = $buildRoot
    $runnerSource = [IO.File]::ReadAllText((Join-Path $projectRoot 'tools/windows/process_job.cs'))
    $runnerSource = $runnerSource.Replace('internal static class OwnedCommand', 'public static class OwnedCommand').Replace('internal static int Run', 'public static int Run')
    if (-not ('OwnedCommand' -as [type])) { Add-Type -TypeDefinition $runnerSource }
    if (-not ('LiXInstallerFileSecurity' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
public static class LiXInstallerFileSecurity
{
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    private static extern uint GetNamedSecurityInfo(string path, int type, uint information,
        IntPtr owner, IntPtr group, IntPtr dacl, out IntPtr sacl, out IntPtr descriptor);
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool GetAce(IntPtr acl, uint index, out IntPtr ace);
    [DllImport("advapi32.dll")]
    private static extern IntPtr GetSidSubAuthorityCount(IntPtr sid);
    [DllImport("advapi32.dll")]
    private static extern IntPtr GetSidSubAuthority(IntPtr sid, uint index);
    [DllImport("kernel32.dll")]
    private static extern IntPtr LocalFree(IntPtr memory);
    public static void Check(string path)
    {
        IntPtr sacl, descriptor;
        uint status = GetNamedSecurityInfo(path, 1, 0x10, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, out sacl, out descriptor);
        if (status != 0) throw new Win32Exception((int)status);
        try
        {
            if (sacl == IntPtr.Zero) return;
            int count = (ushort)Marshal.ReadInt16(sacl, 4);
            for (uint index = 0; index < count; index++)
            {
                IntPtr ace;
                if (!GetAce(sacl, index, out ace)) throw new Win32Exception();
                if (Marshal.ReadByte(ace) != 0x11) continue;
                IntPtr sid = IntPtr.Add(ace, 8);
                byte parts = Marshal.ReadByte(GetSidSubAuthorityCount(sid));
                if (parts == 0 || Marshal.ReadInt32(GetSidSubAuthority(sid, (uint)(parts - 1))) < 8192)
                    throw new InvalidOperationException("Build artifact inherited a low integrity label; use a normal local build directory");
            }
        }
        finally { if (descriptor != IntPtr.Zero) LocalFree(descriptor); }
    }
}
'@
    }
    [LiXInstallerFileSecurity]::Check($buildRoot)
    $uv = (Get-Command uv -CommandType Application -ErrorAction Stop).Source
    Invoke-OwnedBuild $uv @('run','--offline','python','-m','tools.build_windows_installer','--stage-only',
        '--root',$projectRoot,'--payload',$Payload,'--uv-archive',$UvArchive,'--license-root',$LicenseRoot) 120
    $manifest = [IO.File]::ReadAllText((Join-Path $Payload 'LiX.install.json')) | ConvertFrom-Json
    $stage = [IO.File]::ReadAllText((Join-Path $Payload 'STAGE.json')) | ConvertFrom-Json
    if ($manifest.version -notmatch '^\d+\.\d+\.\d+$' -or $manifest.version -ne $stage.version) { throw 'Invalid staged version' }
    $setup = Join-Path $buildRoot 'inno-setup.exe'
    [IO.File]::WriteAllBytes($setup, [IO.File]::ReadAllBytes($InnoSetup))
    [LiXInstallerFileSecurity]::Check($setup)
    $toolRoot = Join-Path $buildRoot 'inno'
    Invoke-OwnedBuild $setup @('/PORTABLE=1','/SP-','/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/CURRENTUSER','/NOICONS',('/DIR='+$toolRoot)) 120
    $compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    $launcher = Join-Path $buildRoot 'LiX.exe'
    Invoke-OwnedBuild $compiler @('/nologo','/target:winexe','/platform:x64','/optimize+',
        '/reference:System.Windows.Forms.dll','/reference:System.Drawing.dll','/reference:System.Web.Extensions.dll',
        ('/out:'+$launcher),('/win32icon:'+(Join-Path $projectRoot 'office_agent/webui/static/brand/lix.ico')),
        ('/resource:'+(Join-Path $Payload 'bootstrap.json')+',LiX.Bootstrap'),
        (Join-Path $projectRoot 'tools/windows/installed_launcher.cs'),(Join-Path $projectRoot 'tools/windows/process_job.cs'),
        (Join-Path $Payload 'launcher_version.cs')) 60
    [IO.File]::Move($launcher, (Join-Path $Payload 'LiX.exe'))
    [LiXInstallerFileSecurity]::Check((Join-Path $Payload 'LiX.exe'))
    $iscc = Join-Path $toolRoot 'ISCC.exe'
    [LiXInstallerFileSecurity]::Check($iscc)
    $banner = & $iscc --version
    if ($LASTEXITCODE -ne 0 -or "$banner" -notmatch '\b7\.1\.1\b') { throw 'Expected Inno Setup 7.1.1' }
    $binaryRoot = Join-Path $buildRoot 'output'
    [IO.Directory]::CreateDirectory($binaryRoot) | Out-Null
    Invoke-OwnedBuild $iscc @('/Qp',('/DPayloadDir='+$Payload),('/DOutputDir='+$binaryRoot),
        ('/DVersion='+$manifest.version),(Join-Path $projectRoot 'tools/windows/LiX.iss')) 180
    $name = 'LiX-Setup-' + $manifest.version + '-preview-x64.exe'
    $artifact = Join-Path $binaryRoot $name
    if ((Get-Item -LiteralPath $artifact).Length -gt 128MB) { throw 'Installer exceeds the size bound' }
    $checksum = (Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.Directory]::CreateDirectory($Output) | Out-Null
    [IO.File]::Move($artifact, (Join-Path $Output $name))
    [LiXInstallerFileSecurity]::Check((Join-Path $Output $name))
    [IO.File]::WriteAllText((Join-Path $Output 'SHA256SUMS.txt'), ($checksum+'  '+$name+"`n"), [Text.Encoding]::ASCII)
    $receipt = [ordered]@{file=$name;bytes=(Get-Item -LiteralPath (Join-Path $Output $name)).Length;sha256=$checksum;
        version=$manifest.version;channel='network-preview';signed=$false;published=$false;
        public_source_files=$stage.public_source_files;payload_files=$stage.payload_files;
        source_manifest_sha256=$stage.source_manifest_sha256;binary_created_in_system_temp=$true;
        final_binary_same_volume_move=$true;low_integrity_artifacts_rejected=$true;security_permissions_modified=$false}
    [IO.File]::WriteAllText((Join-Path $Output 'BUILD.json'), (($receipt | ConvertTo-Json)+"`n"), $utf8)
    $receipt | ConvertTo-Json
}
finally {
    $env:TEMP = $originalTemp
    $env:TMP = $originalTmp
    $resolvedBuild = [IO.Path]::GetFullPath($buildRoot)
    if ([IO.Path]::GetDirectoryName($resolvedBuild).Equals($temporaryParent.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedBuild).StartsWith('lix-installer-build-')) {
        Remove-Item -LiteralPath $resolvedBuild -Recurse -Force -ErrorAction SilentlyContinue
    }
}
