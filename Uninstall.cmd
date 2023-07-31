@echo off
setlocal EnableDelayedExpansion

for %%v in (2022) do (
    for %%s in (Preview Enterprise Professional Community) do (
        set "vsixInstaller=%ProgramW6432%\Microsoft Visual Studio\%%v\%%s\Common7\IDE\VSIXInstaller.exe"
        if exist "!vsixInstaller!" (
            echo Found Visual Studio %%v %%s
            "!vsixInstaller!" %* /uninstall:VSFormatOnSave2022.9484441C-1BE1-4481-997F-0FEDF16868C8
            goto :Uninstall2022End
        )
    )
)
:Uninstall2022End

for %%v in (2019 2017) do (
    for %%s in (Preview Enterprise Professional Community) do (
        set "vsixInstaller=%ProgramFiles(x86)%\Microsoft Visual Studio\%%v\%%s\Common7\IDE\VSIXInstaller.exe"
        if exist "!vsixInstaller!" (
            echo Found Visual Studio %%v %%s
            "!vsixInstaller!" %* /uninstall:73e31bd4-5e3f-4160-a212-ac992f0a9126
            exit /b
        )
    )
)

set "vsixInstaller=%ProgramFiles(x86)%\Microsoft Visual Studio 14.0\Common7\IDE\VSIXInstaller.exe"
if exist "%vsixInstaller%" (
    echo Found Visual Studio 2015
    "%vsixInstaller%" %* /uninstall:73e31bd4-5e3f-4160-a212-ac992f0a9126
    exit /b
)

endlocal
