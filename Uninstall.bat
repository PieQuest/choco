@echo off

:: Execute the PS1 file with the same name as this batch file.
set filename=%~d0%~p0%~n0.ps1

if exist "%filename%" (
  PowerShell.exe -ExecutionPolicy Unrestricted -Command "& '%filename%'"
) else (
  echo The file: %filename% is not found.
)

exit