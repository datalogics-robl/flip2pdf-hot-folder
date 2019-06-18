$folder = 'C:\Users\robl\src\BlogArticle\watched_dir'
$destinationFolder = 'C:\Users\robl\src\BlogArticle\destination_dir\'

$officeTypes = ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"
$postscriptTypes = ".ps", ".eps"
$imageTypes = ".bmp", ".jpg", ".jpeg", ".png", ".tif", ".tiff"
$supportedTypes = $officeTypes + $postscriptTypes + $imageTypes

$fsw = New-Object IO.FileSystemWatcher $folder, "*.*" -Property @{
  IncludeSubdirectories = $true
  NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'
}
$onCreated = Register-ObjectEvent $fsw Created -SourceIdentifier FileConverted -Action {
  $path = $Event.SourceEventArgs.FullPath
  $name = $Event.SourceEventArgs.Name
  $timeStamp = $Event.TimeGenerated
  $extension = [System.IO.Path]::GetExtension($name).ToLower()
  $outputName = $destinationFolder + [System.IO.Path]::GetFileNameWithoutExtension($name) + ".pdf"
 
  if ($supportedTypes.Contains($extension)) {
    Write-Host "The file '$name' was converted at $timeStamp"
    Invoke-Command -ScriptBlock { flip2pdf --input $path --output $outputName }
    Move-Item $path -Destination $destinationFolder -Force -Verbose # Force will overwrite files with the same name
  }
}
