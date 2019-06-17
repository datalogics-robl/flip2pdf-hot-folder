$folder = 'C:\Users\robl\src\BlogArticle\watched_dir'
$destinationFolder = 'C:\Users\robl\src\BlogArticle\destination_dir\'
$outputName = $destinationFolder + 'output.pdf'

$supportedTypes = ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".ps", ".eps" # TODO: add image types

$fsw = New-Object IO.FileSystemWatcher $folder, "*.*" -Property @{
  IncludeSubdirectories = $true
  NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'
}
$onCreated = Register-ObjectEvent $fsw Created -SourceIdentifier FileConverted -Action {
  $path = $Event.SourceEventArgs.FullPath
  $name = $Event.SourceEventArgs.Name
  $timeStamp = $Event.TimeGenerated
  $extension = [System.IO.Path]::GetExtension($name).ToLower()
 
  if ($supportedTypes.Contains($extension)) {
    Write-Host "The file '$name' was converted at $timeStamp"
    Invoke-Command -ScriptBlock { flip2pdf --input $path --output $outputName }
    Move-Item $path -Destination $destinationFolder -Force -Verbose # Force will overwrite files with the same name
  }
}
