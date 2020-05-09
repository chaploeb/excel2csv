#Heavily modified from https://gallery.technet.microsoft.com/office/How-to-convert-Excel-xlsx-d9521619

#Recurses through a folder that contains Excel files, converting each one to CSV and moving it as it is converted.  

#Warning, this program is destructive on the original folder in a few ways, which may not be ideal--
#it was what I needed for my use case but may not make sense for other folders.

#In particular, the program as currently written ruthlessly removes printer settings from every excel file it encounters.

$ErrorActionPreference = 'Stop'

function releaseExcel {
    param ($excelApp)
    # Release Excel Com Object resource
    $excelApp.Workbooks.Close()
    $excelApp.Visible = $true
    Start-Sleep 5
    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
}

function deleteEmptyFolders {
    param ($path)
    do {
      $dirs = Get-ChildItem $path -directory -recurse | Where-Object { (Get-ChildItem $_.fullName).count -eq 0 } | Select-Object -expandproperty FullName
      $dirs | Foreach-Object { Remove-Item $_ -Verbose}
    } while ($dirs.count -gt 0)
}


function excelZip {
    # remove printersettings1.bin from Excel file by treating it as a .zip file.  

    param ($zipfile)

    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression')

    $files   = 'printersettings1.bin'

    $stream = New-Object IO.FileStream($zipfile, [IO.FileMode]::Open)
    $mode   = [IO.Compression.ZipArchiveMode]::Update
    $zip    = New-Object IO.Compression.ZipArchive($stream, $mode)

    ($zip.Entries | Where-Object { $files -contains $_.Name }) | ForEach-Object { $_.Delete() }

    $zip.Dispose()
    $stream.Close()
    $stream.Dispose()
}

$FolderPath = "C:\DCL\greenstar_latest"                      #Source folder containing Excel files     
$OutputPath = "C:\DCL\greenstar_latest_csv"                  #Destination folder for CSV files
$ProcessedPath = "C:\DCL\greenstar_latest_processed"         #Excel files are moved here as they are processed

#cleanup

#Clean garbage files
Get-ChildItem -Path $FolderPath -Include ~*.xlsx, ~*.xls -Recurse | Remove-Item -Verbose
Get-ChildItem -Path $FolderPath -Include .DS_Store -Recurse | Remove-Item -Verbose
#Clean empty folders
deleteEmptyFolders ($FolderPath)


$ExcelFiles = Get-ChildItem -Path $FolderPath -Include *.xlsx, *.xls -Recurse

$excelApp = New-Object -ComObject Excel.Application
$excelApp.DisplayAlerts = $false

$ExcelFiles | ForEach-Object {
    Write-Output ($_.FullName)

    # Remove printer settings from Excel file, if present.
    $null=excelZip ($_)

    try {
	    $workbook = $excelApp.Workbooks.Open($_.FullName)
        $workbook.Sheets.Item(1).Columns("F:F").NumberFormat = "#,##0.00"

	    $csvFilePath = ($_.FullName -replace [regex]::Escape($FolderPath), $OutputPath -replace "\.xlsx$", ".csv" -replace "\.xls$", ".csv")
        $null=New-Item -ItemType Directory -Force -Path (Split-Path -path $csvFilePath)
        Write-Output ($csvFilePath)
        $workbook.SaveAs($csvFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
        $workbook.Close()
        Start-Sleep 0.5
        $processFilePath = ($_.FullName -replace [regex]::Escape($FolderPath), $ProcessedPath)
        $null=New-Item -ItemType Directory -Force -Path (Split-Path -path $processFilePath)
        Write-Output ($processFilePath)
        Move-Item  $_.FullName -Destination $processFilePath
    }
    catch {
        Write-Host $_
        ReleaseExcel ($excelApp)
        Exit
    }
}

releaseExcel ($excelApp)
deleteEmptyFolders ($FolderPath)

