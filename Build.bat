@echo off

SET AdditionalsDir=.\Additionals
SET HelperScripts=%AdditionalsDir%\Scripts
SET SetupScripts=%AdditionalsDir%\Setup
SET NuGetDir=%AdditionalsDir%\NuGet
SET ProjectPaths='.\DryIoc\DryIoc.Excel\PrismTaskPanes.DryIoc.Excel.csproj','.\DryIoc\DryIoc.PowerPoint\PrismTaskPanes.DryIoc.PowerPoint.csproj','.\Base\PrismTaskPanes.Commons\PrismTaskPanes.Commons.csproj'

echo.
echo ##### Create PrismTaskPanes #####
echo.

CHOICE /C dr /N /M "Shall the [d]ebug version or the [r]elease version be compiled?"
SET CONFIGSELECTION=%ERRORLEVEL%
echo.

if /i "%ERRORLEVEL%" == "2" GOTO RELEASE

:DEBUG

SET CONFIGURATION=Debug

echo.
echo Clean solution
echo.

CALL "%HelperScripts%\Unregister.bat"

GOTO BUILD

:RELEASE

SET CONFIGURATION=Release

CHOICE /C mb /N /M "Shall the [b]uild (x.x._X_.0) or the [m]inor version (x._X_.0.0) be increased?"
SET VERSIONSELECTION=%ERRORLEVEL%
echo.

if /i "%VERSIONSELECTION%" == "1" (
	echo.
	echo Update minor version
	echo.

	powershell "%SetupScripts%\Update_VersionMinor.ps1 -projectPaths %ProjectPaths%"
)

echo.
echo Clean solution
echo.

CALL "%HelperScripts%\Unregister.bat"
CALL "%HelperScripts%\Clean.bat"

GOTO BUILD

:BUILD

echo.
echo Build solution
echo.

dotnet build ".\DryIoc\DryIoc.Excel\PrismTaskPanes.DryIoc.Excel.csproj" --configuration %CONFIGURATION%
dotnet build ".\DryIoc\DryIoc.PowerPoint\PrismTaskPanes.DryIoc.PowerPoint.csproj" --configuration %CONFIGURATION%
dotnet build ".\Base\PrismTaskPanes.Commons\PrismTaskPanes.Commons.csproj" --configuration %CONFIGURATION%

if %CONFIGURATION% == Debug (

	dotnet build ".\Examples\DryIoc\Excel\ExcelAddIn1\ExcelAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build ".\Examples\DryIoc\Excel\ExcelAddIn2\ExcelAddIn2.csproj" --configuration %CONFIGURATION%

	dotnet build ".\Examples\DryIoc\PowerPoint\PowerPointAddIn1\PowerPointAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build ".\Examples\DryIoc\PowerPoint\PowerPointAddIn2\PowerPointAddIn2.csproj" --configuration %CONFIGURATION%

	echo.
	CALL "%HelperScripts%\Register.bat"
	
	echo.
	PAUSE

	START excel.exe
	START powerpnt.exe

) else (

	echo.
	echo Copy NuGet packages
	echo.

	del %NuGetDir%\*.nupkg
	for /R %cd% %%f in (*.nupkg) do copy %%f %NuGetDir%\

	echo.
	echo Update build version
	echo.

	powershell "%SetupScripts%\Update_VersionBuild.ps1 -projectPaths %ProjectPaths%"

	echo.
	PAUSE
)