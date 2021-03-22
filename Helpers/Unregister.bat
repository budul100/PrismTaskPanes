@ECHO off

SET ExampleDirectory=..\Examples

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister %ExampleDirectory%\Excel\ExcelAddIn1\bin\Debug\net472\ExcelAddIn1.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister %ExampleDirectory%\Excel\ExcelAddIn1\bin\Debug\net472\PrismTaskPanes.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister %ExampleDirectory%\Excel\ExcelAddIn2\bin\Debug\net472\ExcelAddIn2.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister %ExampleDirectory%\Excel\ExcelAddIn2\bin\Debug\net472\PrismTaskPanes.dll