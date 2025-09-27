@echo off
echo ================================
echo ReportGenerator Release Helper
echo ================================
echo.

if "%1"=="" (
    echo Usage: release.bat v1.0.0
    echo Example: release.bat v1.2.3
    echo.
    echo This will:
    echo 1. Commit current changes
    echo 2. Create and push a git tag
    echo 3. Trigger automatic GitHub release
    echo.
    pause
    exit /b 1
)

set VERSION=%1

echo Creating release version: %VERSION%
echo.

echo Step 1: Adding all changes...
git add .
if errorlevel 1 (
    echo Error: Failed to add changes
    pause
    exit /b 1
)

echo Step 2: Committing changes...
set /p COMMIT_MSG="Enter commit message (or press Enter for default): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=Release %VERSION%

git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo Warning: No changes to commit or commit failed
)

echo Step 3: Creating tag %VERSION%...
git tag %VERSION%
if errorlevel 1 (
    echo Error: Failed to create tag
    pause
    exit /b 1
)

echo Step 4: Pushing tag to GitHub...
git push origin %VERSION%
if errorlevel 1 (
    echo Error: Failed to push tag
    pause
    exit /b 1
)

echo.
echo ================================
echo SUCCESS! 
echo ================================
echo.
echo Your release is being built automatically.
echo.
echo Check progress at:
echo https://github.com/adklinge1/ReportGenerations/actions
echo.
echo Release will be available at:
echo https://github.com/adklinge1/ReportGenerations/releases
echo.
echo This usually takes 2-3 minutes to complete.
echo.
pause