PUSHD ..

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister .\Tests\TestAddIn1\bin\Debug\net472\TestAddIn1.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister .\Tests\TestAddIn1\bin\Debug\net472\PrismTaskPanes.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister .\Tests\TestAddIn2\bin\Debug\net472\TestAddIn2.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister .\Tests\TestAddIn2\bin\Debug\net472\PrismTaskPanes.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister .\Tests\TestAddIn1\bin\Release\net472\TestAddIn1.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister .\Tests\TestAddIn1\bin\Release\net472\PrismTaskPanes.dll

POPD