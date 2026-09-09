@echo off
setlocal
set "import_exit_code=1"
cd /d "%~dp0"
echo Opdaterer regnskabet...
echo.
where python >nul 2>nul
if not errorlevel 1 goto use_python
goto use_py
:missing_python
echo Python mangler. Installer Python 3.10 eller nyere, og start igen.
goto finish
:use_py
where py >nul 2>nul
if errorlevel 1 goto missing_python
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
if errorlevel 1 goto old_python
py -3 -I "program\importer.py"
set "import_exit_code=%errorlevel%"
goto finish
:use_python
python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
if errorlevel 1 goto use_py
python -I "program\importer.py"
set "import_exit_code=%errorlevel%"
goto finish
:old_python
echo Python 3.10 eller nyere er noedvendig. Opdater Python og start igen.
:finish
echo.
pause
exit /b %import_exit_code%
