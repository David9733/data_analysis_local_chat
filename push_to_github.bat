@echo off
cd /d "%~dp0"
set "PATH=%USERPROFILE%\.local\mingit\cmd;%PATH%"
echo Pushing to GitHub (David9733/data_analysis_local_chat)...
git push -u origin main
echo.
pause
