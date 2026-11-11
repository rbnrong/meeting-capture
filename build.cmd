@echo off
setlocal
set "MC_ROOT=%~dp0"
set "MC_CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%MC_CSC%" set "MC_CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%MC_CSC%" (
  echo Windows .NET Framework compiler not found.
  exit /b 1
)
if not exist "%MC_ROOT%dist" mkdir "%MC_ROOT%dist"
"%MC_CSC%" /nologo /target:winexe /platform:x64 /optimize+ /codepage:65001 /warnaserror+ /win32manifest:"%MC_ROOT%app\app.manifest" /win32icon:"%MC_ROOT%app\app.ico" /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /out:"%MC_ROOT%dist\MeetingCapture.exe" "%MC_ROOT%app\MeetingCapture.cs"
if errorlevel 1 exit /b 1
echo Built: %MC_ROOT%dist\MeetingCapture.exe
exit /b 0
