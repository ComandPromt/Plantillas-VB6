@echo off

REM This will create the Resource File
REM RC.EXE is the Resource Compiler in the path VB98\WIZARD\RC.EXE

echo. 
echo Creating Resource File...
echo.

CALL RC.EXE EXAMPLE.RC

echo. 
echo Done.
echo.
