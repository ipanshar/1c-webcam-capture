@echo off
echo ================================================
echo Регистрация COM компонента AsterWebCamTo1C
echo ================================================
echo.

REM Проверка прав администратора
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ОШИБКА: Требуются права администратора!
    echo Запустите этот файл от имени администратора.
    echo.
    pause
    exit /b 1
)

echo Переход в каталог Release...
cd /d "%~dp0bin\Release\net472"

echo.
echo Регистрация для 32-битных приложений (x86)...
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe AsterWebCamTo1C.dll /codebase /tlb:AsterWebCamTo1C.tlb

if %errorLevel% neq 0 (
    echo.
    echo ОШИБКА при регистрации x86!
    pause
    exit /b 1
)

echo.
echo Регистрация для 64-битных приложений (x64)...
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe AsterWebCamTo1C.dll /codebase /tlb:AsterWebCamTo1C.tlb

if %errorLevel% neq 0 (
    echo.
    echo ОШИБКА при регистрации x64!
    pause
    exit /b 1
)

echo.
echo ================================================
echo Регистрация успешно завершена!
echo ================================================
echo.
echo ProgId: AsterWebCamTo1C.WebCameraCapture
echo Namespace: AsterWebCamTo1C
echo.
echo Для использования в 1С:
echo   КамераOLE = Новый COMОбъект("AsterWebCamTo1C.WebCameraCapture");
echo.
pause
