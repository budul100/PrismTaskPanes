param (
	[string] $baseDir
)

$baseDir = if ([System.IO.Path]::IsPathRooted("$baseDir\")) {"$baseDir\"} else {[System.IO.Path]::GetFullPath((Join-Path $pwd "$baseDir\"))} 

Write-Host "All subfolders on $baseDir cleaned now."

$dirs = Get-ChildItem $baseDir -directory -recurse | Where-Object { (Get-ChildItem $_.fullName).count -eq 0 } | Select-Object -expandproperty FullName
$dirs | Foreach-Object { Remove-Item $_ }

Get-ChildItem $baseDir -include bin,obj,_debug,_Out -Recurse | ForEach-Object ($_) { Remove-Item $_.FullName -Recurse -Force }