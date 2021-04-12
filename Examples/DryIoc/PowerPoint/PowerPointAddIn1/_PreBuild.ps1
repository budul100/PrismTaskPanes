function RemoveKey([string]$root, [Microsoft.PowerShell.Commands.MatchInfo]$contents)
{
    $contents.matches | ForEach-Object { 
        $key="${root}:\SOFTWARE\Classes\$_" 
        Write-Host "Check $key."
        if (Test-Path -Path $key)
        {
            Write-Host "Delete $key." 
            Remove-Item $key -Recurse
        }
    }       
}

Get-ChildItem -Path .\ -Filter *.cs | ForEach-Object {
    $progId = Get-Content $_ | Select-String -Pattern '(?<=ProgId\(")[^"]+'
    
    RemoveKey "HKLM" $progId
    RemoveKey "HKCU" $progId

    $clsId = Get-Content $_ | Select-String -Pattern '(?<=Guid\(")[^"]+'
    
    RemoveKey "HKLM" $clsId
    RemoveKey "HKCU" $clsId
}