[Setup]
AppId={{F994C926-AD9B-4F77-85CD-9BB282C05CB7}
AppName=LiX
AppVersion={#Version}
#ifdef Offline
AppVerName=LiX {#Version} 离线内部预览版
ExtraDiskSpaceRequired={#ExtraSpace}
#else
AppVerName=LiX {#Version} 联网预览版
#endif
AppPublisher=LiX contributors
AppPublisherURL=https://github.com/liwiki03-hue/LiX-OfficeAgent
AppSupportURL=https://github.com/liwiki03-hue/LiX-OfficeAgent/issues
DefaultDirName={localappdata}\Programs\LiX
DefaultGroupName=LiX
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
WizardStyle=modern dynamic
WizardSizePercent=110
DisableProgramGroupPage=yes
DisableDirPage=no
LicenseFile={#PayloadDir}\app\LICENSE
InfoBeforeFile={#PayloadDir}\INSTALL-NOTES.txt
SetupIconFile={#PayloadDir}\app\office_agent\webui\static\brand\lix.ico
UninstallDisplayIcon={app}\LiX.exe
OutputDir={#OutputDir}
#ifndef Flavor
#define Flavor "preview"
#endif
OutputBaseFilename=LiX-Setup-{#Version}-{#Flavor}-x64
Compression=lzma2
SolidCompression=yes
AppMutex=LiXInstalledRuntime
CloseApplications=no
RestartApplications=no
SetupLogging=yes
UninstallLogging=yes
VersionInfoDescription=LiX Windows installer preview

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："; Flags: unchecked

[Files]
#include PayloadDir + "\FILES.iss"

[Icons]
Name: "{group}\LiX"; Filename: "{app}\LiX.exe"; WorkingDir: "{app}"
Name: "{group}\卸载 LiX"; Filename: "{uninstallexe}"
Name: "{autodesktop}\LiX"; Filename: "{app}\LiX.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
#ifdef Offline
Filename: "{app}\LiX.exe"; Description: "启动 LiX（本地组件装配）"; Flags: nowait postinstall skipifsilent
#else
Filename: "{app}\LiX.exe"; Description: "启动 LiX（首次使用需联网准备）"; Flags: nowait postinstall skipifsilent
#endif

[Code]
function GetFileAttributesW(lpFileName: string): LongWord;
  external 'GetFileAttributesW@kernel32.dll stdcall';

function SafeDirectory(Value: string): Boolean;
var
  Parent: string;
  Attributes: LongWord;
begin
  Result := False;
  if (Pos('\\', Value) = 1) or (Length(Value) < 4) then Exit;
  while Length(Value) > 3 do begin
    Attributes := GetFileAttributesW(Value);
    if (Attributes <> $FFFFFFFF) and ((Attributes and $400) <> 0) then Exit;
    Parent := ExtractFileDir(Value);
    if Parent = Value then Exit;
    Value := Parent;
  end;
  Result := True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Target: string;
begin
  Result := True;
  if CurPageID = wpSelectDir then begin
    Target := ExpandFileName(WizardDirValue);
    if not SafeDirectory(Target) then begin
      MsgBox('请选择本机普通文件夹，不能使用链接或重定向目录。', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

#ifdef Offline
function WebViewPresent(): Boolean;
var
  Value: string;
begin
  Result := (RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Value) or
    RegQueryStringValue(HKCU32, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Value)) and
    (Value <> '') and (Value <> '0.0.0.0');
end;

function InstallPrerequisite(Name, Parameters: string; Elevate: Boolean; var NeedsRestart: Boolean): string;
var
  FileName: string;
  Code: Integer;
  Started: Boolean;
begin
  Result := '';
  ExtractTemporaryFile(Name + '.exe');
  FileName := ExpandConstant('{tmp}\') + Name + '.exe';
  if Elevate then
    Started := ShellExec('runas', FileName, Parameters, '', SW_SHOW, ewWaitUntilTerminated, Code)
  else
    Started := Exec(FileName, Parameters, '', SW_SHOW, ewWaitUntilTerminated, Code);
  if not Started then begin
    Result := '先决组件未能启动或已取消：' + Name;
    Exit;
  end;
  if (Code = 3010) or (Code = 1641) then begin
    NeedsRestart := True;
    Result := '先决组件需要重启 Windows；重启后运行同一离线安装包。';
    Exit;
  end;
  if Code <> 0 then Result := '先决组件安装未完成：' + Name + '，退出码 ' + IntToStr(Code);
end;
#endif

function PrepareToInstall(var NeedsRestart: Boolean): string;
var
  Target: string;
begin
  Result := '';
  Target := ExpandConstant('{app}');
  if not SafeDirectory(Target) then begin
    Result := '安装位置已改变或包含链接；未安装。';
    Exit;
  end;
#ifdef Offline
  if FileExists(Target + '\LiX.install.json') and
    (CompareText(GetSHA256OfFile(Target + '\LiX.install.json'), '{#ManifestHash}') <> 0) then begin
    Result := '此内部候选尚不支持跨构建覆盖升级。请保留原安装和数据，选择独立目录。';
    Exit;
  end;
  if not IsDotNetInstalled(net48, 0) then begin
    if WizardSilent then begin Result := '缺少.NET4.8；请使用交互安装允许官方离线组件安装。'; Exit; end;
    Result := InstallPrerequisite('net48', '/passive /norestart', True, NeedsRestart);
    if (Result <> '') or not IsDotNetInstalled(net48, 0) then begin
      if Result = '' then Result := '.NET4.8尚未就绪，请重启后重试。';
      Exit;
    end;
  end;
  if not WebViewPresent() then begin
    Result := InstallPrerequisite('webview2', '/silent /install', False, NeedsRestart);
    if (Result <> '') or not WebViewPresent() then begin
      if Result = '' then Result := 'WebView2 Runtime尚未就绪，请检查官方安装结果。';
      Exit;
    end;
  end;
#else
  if not IsDotNetInstalled(net48, 0) then begin
    Result := 'LiX 需要 .NET Framework 4.8。请先通过 Windows Update 准备，再重新安装。';
    Exit;
  end;
#endif
  if DirExists(Target) and not FileExists(Target + '\LiX.install.json') then begin
    Result := '该文件夹已经存在且不是本安装器建立的 LiX 目录，请选择新的文件夹。';
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigFile: string;
begin
  if CurStep = ssPostInstall then begin
    ConfigFile := ExpandConstant('{app}\app\.env');
    if not FileExists(ConfigFile) then begin
      if not SaveStringToFile(ConfigFile, '', False) then
        RaiseException('无法创建本机配置；请核对安装位置写入权限。');
    end;
  end;
end;
