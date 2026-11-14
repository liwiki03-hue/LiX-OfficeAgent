[Setup]
AppId={{F994C926-AD9B-4F77-85CD-9BB282C05CB7}
AppName=LiX
AppVersion={#Version}
AppVerName=LiX {#Version} 联网预览版
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
Filename: "{app}\LiX.exe"; Description: "启动 LiX（首次使用需联网准备）"; Flags: nowait postinstall skipifsilent

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
  if not IsDotNetInstalled(net48, 0) then begin
    Result := 'LiX 需要 .NET Framework 4.8。请先通过 Windows Update 准备，再重新安装。';
    Exit;
  end;
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
