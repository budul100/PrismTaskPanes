@ECHO off

SET ExampleDirectory=..\Examples

REM "%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase  %ExampleDirectory%\Excel\ExcelAddIn1\bin\Debug\net472\PrismTaskPanes.DryIoc.Excel.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase  %ExampleDirectory%\Excel\ExcelAddIn1\bin\Debug\net472\ExcelAddIn1.dll

REM "%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase  %ExampleDirectory%\Excel\ExcelAddIn2\bin\Debug\net472\PrismTaskPanes.DryIoc.Excel.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase  %ExampleDirectory%\Excel\ExcelAddIn2\bin\Debug\net472\ExcelAddIn2.dll