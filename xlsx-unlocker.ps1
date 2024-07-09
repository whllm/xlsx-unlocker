# The following powershell script iterates through individual sheets and removes protection 
# Protection is only removed at the sheet level.
# This does nothing for documents which are password protected at the document level.

param (
    [string[]]$InputPath
)

# Get the current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# If no input path is provided, prompt the user to select one or more files using a file dialog
if (-not $InputPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $OpenFileDialog.InitialDirectory = $scriptDir  # Set the initial directory to script directory
    $OpenFileDialog.Multiselect = $true  # Allow multiple file selection

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $InputPath = $OpenFileDialog.FileNames
    } else {
        Write-Error "No files selected. Exiting script."
        exit
    }
}

foreach ($xlsxPath in $InputPath) {
    # Ensure the input path is an absolute path
    $xlsxPath = (Resolve-Path -Path $xlsxPath).Path

    # Extract just the filename from the input path
    $xlsxName = Split-Path -Leaf $xlsxPath

    # Define the temporary directory for extraction
    $tempDir = Join-Path $scriptDir "temp_$($xlsxName -replace '\.xlsx$')"

    # Ensure the temporary directory exists
    if (Test-Path $tempDir) {
        Remove-Item -Recurse -Force $tempDir
    }
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    # Open the XLSX as a zip archive and extract all contents
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($xlsxPath, $tempDir)

    # Get all sheetX.xml files
    $sheetFiles = Get-ChildItem -Path "$tempDir\xl\worksheets" -Filter "sheet*.xml"

    # Define the pattern to match <sheetpro(.*?)\/>
    $pattern = "<sheetpro(.*?)\/>"

    foreach ($sheetFile in $sheetFiles) {
        # Read the contents of the sheetX.xml file
        $sheetContent = Get-Content $sheetFile.FullName

        # Use a regex to remove the line matching <sheetpro(.*?)\/>
        $modifiedContent = $sheetContent -replace $pattern, ""

        # Save the modified content back to the file
        Set-Content -Path $sheetFile.FullName -Value $modifiedContent
    }

    # Create a new ZIP archive with the modified files
    $modifiedZipPath = Join-Path $scriptDir "unlocked_$xlsxName"
    if (Test-Path $modifiedZipPath) {
        Remove-Item $modifiedZipPath
    }

    Add-Type -AssemblyName System.IO.Compression
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $modifiedZipPath)

    # Clean up temporary files and directories
    Remove-Item -Recurse -Force $tempDir

    Write-Output "Modified file created at: $modifiedZipPath"
}
