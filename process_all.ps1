# Get all Excel files in the sinteticke folder
$files = Get-ChildItem -Path "E:\ZavrsniRad\sintetičke" -Filter "*.xlsx"

# Define the output directory
$outputBaseDir = "E:\ZavrsniRad\excel_to_csv_dcat\output_all_v2"

# Process each file
foreach ($file in $files) {
    Write-Host "Processing $($file.Name)..."
    # Create a subdirectory for each Excel file's output
    $fileSpecificOutputDir = Join-Path -Path $outputBaseDir -ChildPath $file.BaseName
    New-Item -ItemType Directory -Path $fileSpecificOutputDir -ErrorAction SilentlyContinue | Out-Null

    # Construct the command
    $command = "e:\ZavrsniRad\.venv\Scripts\python.exe -m excel_to_csv_dcat `"$($file.FullName)`" -o `"$fileSpecificOutputDir`" --metadata-format turtle"
    Write-Host "Executing: $command"
    Invoke-Expression $command
    Write-Host "Completed processing $($file.Name)`n"
}
