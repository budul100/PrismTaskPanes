SET CONFIGURATION=Debug

CALL Unregister.bat
CALL Clean.bat

dotnet build "..\DryIoc\DryIoc.Excel\PrismTaskPane.DryIoc.Excel.csproj" --configuration %CONFIGURATION%
dotnet build "..\DryIoc\DryIoc.PowerPoint\PrismTaskPane.DryIoc.PowerPoint.csproj" --configuration %CONFIGURATION%

dotnet build "..\Base\PrismTaskPanes.Controls\PrismTaskPanes.Controls.csproj" --configuration %CONFIGURATION%
dotnet build "..\Base\PrismTaskPanes.Regions\PrismTaskPanes.Regions.csproj" --configuration %CONFIGURATION%

if %CONFIGURATION% == Debug (
	dotnet build "..\Examples\DryIoc\Excel\ExcelAddIn1\ExcelAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build "..\Examples\DryIoc\Excel\ExcelAddIn2\ExcelAddIn2.csproj" --configuration %CONFIGURATION%

	dotnet build "..\Examples\DryIoc\PowerPoint\PowerPointAddIn1\PowerPointAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build "..\Examples\DryIoc\PowerPoint\PowerPointAddIn2\PowerPointAddIn2.csproj" --configuration %CONFIGURATION%
)

CALL Register.bat

PAUSE

if %CONFIGURATION% == Debug (
	START excel.exe
	START powerpnt.exe
)