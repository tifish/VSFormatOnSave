@echo off
setlocal EnableDelayedExpansion

for %%v in (2022) do (
    for %%s in (Preview Enterprise Professional Community) do (
        set "vsixInstaller=%ProgramW6432%\Microsoft Visual Studio\%%v\%%s\Common7\IDE\VSIXInstaller.exe"
        if exist "!vsixInstaller!" (
            echo Found Visual Studio %%v %%s
            "!vsixInstaller!" %* /admin "%~dp0VSFormatOnSaveFor2022.vsix"
            goto :Install2022End
        )
    )
)
:Install2022End

for %%v in (2019 2017) do (
    for %%s in (Preview Enterprise Professional Community) do (
        set "vsixInstaller=%ProgramFiles(x86)%\Microsoft Visual Studio\%%v\%%s\Common7\IDE\VSIXInstaller.exe"
        if exist "!vsixInstaller!" (
            echo Found Visual Studio %%v %%s
            "!vsixInstaller!" %* /admin "%~dp0VSFormatOnSave.vsix"
            exit /b
        )
    )
)

set "vsixInstaller=%ProgramFiles(x86)%\Microsoft Visual Studio 14.0\Common7\IDE\VSIXInstaller.exe"
if exist "%vsixInstaller%" (
    echo Found Visual Studio 2015
    "%vsixInstaller%" %* /admin "%~dp0VSFormatOnSave.vsix"
    exit /b
)

endlocal
