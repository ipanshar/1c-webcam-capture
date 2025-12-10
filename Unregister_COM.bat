@echo off
echo ================================================
echo Удаление регистрации COM компонента AsterWebCamTo1C
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
echo Удаление регистрации для 32-битных приложений (x86)...
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe AsterWebCamTo1C.dll /unregister

if %errorLevel% neq 0 (
    echo.
    echo Предупреждение: ошибка при удалении регистрации x86
)

echo.
echo Удаление регистрации для 64-битных приложений (x64)...
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe AsterWebCamTo1C.dll /unregister

if %errorLevel% neq 0 (
    echo.
    echo Предупреждение: ошибка при удалении регистрации x64
)

echo.
echo ================================================
echo Удаление регистрации завершено!
echo ================================================
echo.
pause
