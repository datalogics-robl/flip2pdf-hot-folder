$folder = 'C:\Users\robl\src\BlogArticle\watched_dir'
$destinationFolder = 'C:\Users\robl\src\BlogArticle\destination_dir\'
$logpath = $destinationFolder + "convertfile.log"

# Function to write to our log file
function Write-Log {
  param($message)
  "$(Get-Date -Format G) : $message" | Out-File -FilePath $logpath -Append -Force
}

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
  $extension = [System.IO.Path]::GetExtension($name).ToLower()
  $outputName = $destinationFolder + [System.IO.Path]::GetFileNameWithoutExtension($name) + ".pdf"
 
  if ($supportedTypes.Contains($extension)) {
    Invoke-Command -ScriptBlock { flip2pdf --input $path --output $outputName }
    Move-Item $path -Destination $destinationFolder -Force # Force will overwrite files with the same name

    if ($LASTEXITCODE -eq 0) {
      Write-Log "$name successfully converted to $outputName"
    } else {
      Write-Log "Conversion of $name failed with error code: $LASTEXITCODE"
    }
  }
}
