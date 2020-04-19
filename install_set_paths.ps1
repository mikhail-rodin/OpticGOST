$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$documentsFolder = [Environment]::GetFolderPath("MyDocuments")
$macrosFolder = $documentsFolder + '\Zemax\Macros'
$configDir = '"' + $dir + '\config\' + '"'
Set-Location -Path $dir -PassThru
(gc source\reportsExportTemplate.zpl) -replace '#configpath#', $configDir | Out-File -encoding ASCII analysis_export.zpl
Copy-Item source\analysis_export.zpl -Destination $macrosFolder