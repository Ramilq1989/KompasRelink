@echo off
rem Сборка KompasRelink.exe и библиотеки KompasRelinkLib.dll компилятором .NET Framework 4
cd /d "%~dp0"
set CSC="%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
set REFS=/r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll
set SRC=src\KompasRelink.cs src\KompasLibrary.cs
%CSC% /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 /out:KompasRelink.exe ^
  /win32icon:src\KompasRelink.ico /resource:src\KompasRelink.ico,KompasRelink.ico %REFS% %SRC%
%CSC% /nologo /target:library /platform:anycpu /optimize+ /codepage:65001 /out:KompasRelinkLib.dll ^
  /resource:src\KompasRelink.ico,KompasRelink.ico %REFS% %SRC%
