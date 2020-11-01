@ECHO off

SET ExampleDirectory=..\Examples

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectory%\ExampleAddIn1\bin\Debug\net472\PrismTaskPanes.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectory%\ExampleAddIn1\bin\Debug\net472\ExampleAddIn1.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectory%\ExampleAddIn2\bin\Debug\net472\PrismTaskPanes.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectory%\ExampleAddIn2\bin\Debug\net472\ExampleAddIn2.dll