@echo off

REM setting up directories

set pythonPath=..\python.exe
set laapsangScriptPath=..\laapsang.py

REM Execute laapsang.py

%pythonPath% %laapsangScriptPath%

echo Scripts have been executed and the Program has ended.
pause

























