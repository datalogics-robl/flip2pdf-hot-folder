param (
  [Parameter(Mandatory=$true)]$source,
  [Parameter(Mandatory=$true)]$destination,
  $logpath = $destination + "\convertfile.log"
)

$destination += "\"

# Function to write to our log file
function Write-Log {
  param($message)
  "$(Get-Date -Format G) : $message" | Out-File -FilePath $logpath -Append -Force
}

$fsw = New-Object IO.FileSystemWatcher $source, "*" -Property @{
  IncludeSubdirectories = $true
  NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'
}

$onCreated = Register-ObjectEvent $fsw Created -SourceIdentifier FileConverted -Action {
  $path = $Event.SourceEventArgs.FullPath
  $name = $Event.SourceEventArgs.Name
  $extension = [System.IO.Path]::GetExtension($name).ToLower()
  $outputName = $destination + [System.IO.Path]::GetFileNameWithoutExtension($name) + ".pdf"
 
  # Ignore lock files
  if ($name.StartsWith(".~lock.")) {
    return
  }

  $result = Invoke-Command -ScriptBlock { flip2pdf --input $path --output $outputName }
  Move-Item $path -Destination $destination -Force # Force will overwrite files with the same name

  # Log results
  if ($LASTEXITCODE -eq 0) {
      Write-Log "$name successfully converted to $outputName"
  } else {
      Write-Log "Conversion of $name failed with error code: $LASTEXITCODE, message: $($result[5])"
  }
}
