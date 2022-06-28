@ECHO off

SET ExamplesDir=%~dp0..\..\Examples\DryIoc

regsvr32 "%ExamplesDir%\Excel\ExcelAddIn1\bin\Debug\PrismTaskPanes.Host.comhost.dll"

regsvr32 "%ExamplesDir%\Excel\ExcelAddIn1\bin\Debug\ExcelAddIn1.comhost.dll"
regsvr32 "%ExamplesDir%\Excel\ExcelAddIn2\bin\Debug\ExcelAddIn2.comhost.dll"

regsvr32 "%ExamplesDir%\PowerPoint\PowerPointAddIn1\bin\Debug\PowerPointAddIn1.comhost.dll"
regsvr32 "%ExamplesDir%\PowerPoint\PowerPointAddIn2\bin\Debug\PowerPointAddIn2.comhost.dll"