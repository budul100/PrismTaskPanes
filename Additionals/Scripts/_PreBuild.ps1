param([string]$baseDir)

function RemoveProgIdKey([string]$root, [Microsoft.PowerShell.Commands.MatchInfo]$contents)
{
    $contents.matches | ForEach-Object {
        if ($_)
        {
            $key="${root}:\SOFTWARE\Classes\$_" 
            Write-Host "Check $key."
            if (Test-Path -Path $key)
            {
                Write-Host "Delete $key." 
                Remove-Item $key -Recurse
            }
        } 
    }       
}

function RemoveClsIdKey([string]$root, [Microsoft.PowerShell.Commands.MatchInfo]$contents)
{
    $contents.matches | ForEach-Object { 
        if ($_)
        {        
            $key="${root}:\SOFTWARE\Wow6432Node\Classes\CLSID\{$_}" 
            Write-Host "Check $key."
            if (Test-Path -Path $key)
            {
                Write-Host "Delete $key." 
                Remove-Item $key -Recurse
            }
        }
    }       
}

Write-Host "All subfolders on $baseDir are checked for registry keys."

Get-ChildItem -Path "${baseDir}\" -Filter *.cs -Recurse | ForEach-Object {
    $progId = Get-Content $_.FullName | Select-String -Pattern '(?<=ProgId\(")[^"]+'
    
    RemoveProgIdKey "HKLM" $progId
    RemoveProgIdKey "HKCU" $progId

    $clsId = Get-Content $_.FullName | Select-String -Pattern '(?<=Guid\(")[^"]+'

    RemoveClsIdKey "HKLM" $clsId
    RemoveClsIdKey "HKCU" $clsId
}