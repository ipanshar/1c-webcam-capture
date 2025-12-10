; Скрипт установки для AsterWebCamTo1C с проверкой и установкой .NET Framework
; Требует Inno Setup 6.0+

#define MyAppName "AsterWebCamTo1C Component"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ipanshar"
#define MyDllName "AsterWebCamTo1C.dll"
; Имя файла установщика .NET (должен лежать в папке redist рядом со скриптом)
#define DotNetInstallerName "NDP472-KB4054530-x86-x64-AllOS-ENU.exe"

[Setup]
AppId={{87654321-4321-4321-4321-210987654321}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer
OutputBaseFilename=AsterWebCamTo1C_Setup
Compression=lzma
SolidCompression=yes
; Нужны права админа для регистрации COM и установки .NET
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
; Ваши файлы (проверьте путь bin\Release\net472)
Source: "bin\Release\net472\{#MyDllName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net472\DirectShowLib.dll"; DestDir: "{app}"; Flags: ignoreversion
; PDB файл можно добавить, если очень нужно, раскомментировав строку ниже:
; Source: "bin\Release\net472\AsterWebCamTo1C.pdb"; DestDir: "{app}"; Flags: ignoreversion

; Установщик .NET Framework (будет упакован внутрь вашего setup.exe)
Source: "redist\{#DotNetInstallerName}"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
; 1. Регистрация для 32-битных приложений (Важно для обычного клиента 1С!)
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/codebase ""{app}\{#MyDllName}"" /tlb"; WorkingDir: "{app}"; StatusMsg: "Регистрация COM (x32)..."; Flags: runhidden

; 2. Регистрация для 64-битных приложений (На случай если у вас сервер 1С x64)
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/codebase ""{app}\{#MyDllName}"" /tlb"; WorkingDir: "{app}"; StatusMsg: "Регистрация COM (x64)..."; Flags: runhidden; Check: IsWin64

[UninstallRun]
; Удаление регистрации x32
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/unregister ""{app}\{#MyDllName}"""; WorkingDir: "{app}"; StatusMsg: "Отмена регистрации COM (x32)..."; Flags: runhidden

; Удаление регистрации x64
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/unregister ""{app}\{#MyDllName}"""; WorkingDir: "{app}"; StatusMsg: "Отмена регистрации COM (x64)..."; Flags: runhidden; Check: IsWin64

[Code]
// Функция проверки наличия .NET 4.7.2
function IsDotNet472Detected(): boolean;
var
    success: boolean;
    release: cardinal;
begin
    // Проверяем ключ реестра (Release 461808 = .NET 4.7.2)
    success := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', release);
    result := success and (release >= 461808);
end;

// Эта функция запускается до начала установки файлов
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
    ResultCode: Integer;
begin
    // Если .NET не найден
    if not IsDotNet472Detected() then
    begin
        // Распаковываем установщик .NET во временную папку (это происходит автоматически Inno Setup, 
        // но нам нужно убедиться, что мы можем его запустить)
        
        if MsgBox('Для работы программы требуется Microsoft .NET Framework 4.7.2.' + #13#10 +
                  'Он будет установлен сейчас. Это может занять несколько минут.', mbConfirmation, MB_YESNO) = IDYES then
        begin
            // Извлекаем файл установщика .NET из нашего setup.exe во времнюю папку
            ExtractTemporaryFile('{#DotNetInstallerName}');

            // Запускаем установку .NET в тихом режиме (/q /norestart)
            if Exec(ExpandConstant('{tmp}\{#DotNetInstallerName}'), '/q /norestart', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
            begin
                // Установка .NET прошла, продолжаем установку нашей программы
                // Можно добавить проверку ResultCode, чтобы убедиться в успехе
            end
            else
            begin
                Result := 'Не удалось установить .NET Framework. Код ошибки: ' + IntToStr(ResultCode);
            end;
        end
        else
        begin
            Result := 'Установка отменена пользователем, так как требуется .NET Framework.';
        end;
    end;
end;