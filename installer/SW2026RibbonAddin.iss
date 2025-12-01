; SW2026RibbonAddin.iss — uses a local "payload" folder next to this script

#define AppName       "SW2026RibbonAddin"
#define AppVersion    "1.0.0"
#define AppPublisher  "Mehdi"
; FIXED: Added an extra '{' to escape the opening brace. 
; In Inno Setup strings, '{{' becomes a single literal '{'.
#define AddinGuid     "{{B67E2D5A-8C73-4A3E-93B6-1761C1A8C0C5}"   ; must match Addin.cs
#define AddinDll      "SW2026RibbonAddin.dll"

; Payload folder lives next to this script:
#define PayloadDir    AddBackslash(SourcePath) + "payload"
#define MainDllPath   PayloadDir + "\" + AddinDll

#ifexist MainDllPath
  ; OK, we found the DLL in payload
#else
  #error "Main DLL not found: " + MainDllPath + #13#10 + \
         "Build your add-in (Release) and copy " + AddinDll + " to the 'payload' folder next to this .iss"
#endif

[Setup]
AppId={{B839B2C3-6B90-4255-9E4D-9B081B2E6F00}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={pf64}\Mehdi\SW2026RibbonAddin
OutputBaseFilename=SW2026RibbonAddin-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
WizardStyle=modern
UninstallDisplayIcon={app}\{#AddinDll}
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; REQUIRED — your add-in
; Source: "{#MainDllPath}"; DestDir: "{app}"; Flags: ignoreversion

; UPDATED: Copy ALL .dll files from the payload folder to ensure dependencies (like Interops or Newtonsoft) are present.
Source: "{#PayloadDir}\*.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Ensure autoload for the installing user (your ComRegisterFunction also does this)
Root: HKCU; Subkey: "Software\SolidWorks\AddinsStartup\{#AddinGuid}"; ValueType: dword; ValueData: 1; Flags: uninsdeletekey

[Run]
; Register COM with 64-bit RegAsm — triggers your [ComRegisterFunction]
; UPDATED: Added extra outer quotes around the parameters.
; This prevents the "The filename, directory name... syntax is incorrect" error in cmd.exe.
Filename: "{cmd}"; \
  Parameters: "/k """"{code:GetRegAsm64}"" ""{app}\{#AddinDll}"" /codebase /nologo"""""; \
  WorkingDir: "{app}"; \
  StatusMsg: "Registering... (Close the command window to continue)"; Flags: waituntilterminated

[UninstallRun]
; Unregister on uninstall — triggers [ComUnregisterFunction]
; UPDATED: Used /c instead of /k and added outer quotes for safety.
Filename: "{cmd}"; \
  Parameters: "/c """"{code:GetRegAsm64}"" ""{app}\{#AddinDll}"" /u /nologo"""""; \
  WorkingDir: "{app}"; \
  RunOnceId: "Unreg-{#AppName}"; Flags: runhidden waituntilterminated

[Code]
function GetRegAsm64(Param: string): string;
begin
  Result := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe');
end;

function IsDotNet48OrLater(): Boolean;
var Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release)
            and (Release >= 528040);  { .NET 4.8 }
end;

function InitializeSetup(): Boolean;
begin
  if not IsDotNet48OrLater then
  begin
    MsgBox('.NET Framework 4.8 or later is required. Enable it in Windows Features or install it, then run the setup again.',
           mbError, MB_OK);
    Result := False;
  end
  else
    Result := True;
end;