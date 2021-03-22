REM @echo off

SET HelpersDir=.\Helpers
SET SetupScripts=%HelpersDir%\Scripts
SET SlnPaths='.\DryIoc\DryIoc.Excel\PrismTaskPane.DryIoc.Excel.csproj','.\DryIoc\DryIoc.PowerPoint\PrismTaskPane.DryIoc.PowerPoint.csproj','.\Base\PrismTaskPanes.Controls\PrismTaskPanes.Controls.csproj','.\Base\PrismTaskPanes.Regions\PrismTaskPanes.Regions.csproj'

echo.
echo Clean solution
echo.

CALL "%HelpersDir%\Unregister.bat"
CALL "%HelpersDir%\Clean.bat"

echo.
echo Build solution
echo.

set "config=d"
set /p "config=Shall the [d]ebug version or the [r]elease version be compiled? "
echo.

if /i "%config%" == "d" (

	SET CONFIGURATION=Debug

) else (

	SET CONFIGURATION=Release

	set "update=m"
	set /p "update=Shall the [m]inor version (x._X_.0.0) or the [b]uild (x.x._X_.0) be increased? "
	echo.

	if /i "%update%" == "m" (
		powershell "%SetupScripts%\Update_VersionMinor.ps1 -projectPaths %SlnPaths%"
	)
)

dotnet build "DryIoc\DryIoc.Excel\PrismTaskPane.DryIoc.Excel.csproj" --configuration %CONFIGURATION%
dotnet build "DryIoc\DryIoc.PowerPoint\PrismTaskPane.DryIoc.PowerPoint.csproj" --configuration %CONFIGURATION%

dotnet build "Base\PrismTaskPanes.Controls\PrismTaskPanes.Controls.csproj" --configuration %CONFIGURATION%
dotnet build "Base\PrismTaskPanes.Regions\PrismTaskPanes.Regions.csproj" --configuration %CONFIGURATION%

if %CONFIGURATION% == Debug (

	dotnet build "Examples\DryIoc\Excel\ExcelAddIn1\ExcelAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build "Examples\DryIoc\Excel\ExcelAddIn2\ExcelAddIn2.csproj" --configuration %CONFIGURATION%

	dotnet build "Examples\DryIoc\PowerPoint\PowerPointAddIn1\PowerPointAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build "Examples\DryIoc\PowerPoint\PowerPointAddIn2\PowerPointAddIn2.csproj" --configuration %CONFIGURATION%

	echo.
	CALL "%HelpersDir%\Register.bat"
	
	echo.
	PAUSE

	START excel.exe
	START powerpnt.exe

) else (

	if /i "%update%" == "b" (
		powershell "%SetupScripts%\Update_VersionBuild.ps1 -projectPaths %SlnPaths%"
	)

	echo.
	PAUSE
)