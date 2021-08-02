Param(
[Parameter(Mandatory=$true)]$file,
[Parameter(Mandatory=$true)]$module
)

$pe = [PowerExcel]::new($file, $true)
. "$(Get-Location)\AutomationMaker.ps1" -action "run" -event $pe -module $module


class PowerExcel {
    [System.IO.FileSystemInfo] $file
    $meta = @{}
    $excelConn = @{
        conn=$null
        workbook=$null
        ws=$null
    }

    PowerExcel($excelFile, $showWindow) {
        $this.file = $excelFile
        $import = (Import-Clixml -Path "$(Get-Location)\savedConfig.xml") #[$mDataKey]
        $this.meta = $import[($import.Keys | Where-Object {$_ -eq $module})]

        #if not an excel or processable file, then return
        #does not need a workbook connection
        if ($excelFile.Extension -ne '.csv' -xor $excelFile.Extension -ne '.xls' -xor $excelFile.Extension -ne '.xlsx') {
            #run extension specific function
            return
        }

        $this.excelConn.conn = (New-Object -ComObject Excel.Application)
        $this.excelConn.workbook = ($this.excelConn.conn.Workbooks.Open($excelFile))
        #save file as XLSX file version 51: xlOpenXMLWorkbook
        if ($excelFile.Extension -eq '.csv') {
            $newFile = "$($this.meta.inProgress)\$($excelFile.BaseName)"
            if ($this.meta.saveToXl) {
                Start-Sleep -Milliseconds 500
                Write-Host "[$newFile] : File name"
            $this.excelConn.workbook.SaveAs($newFile, 51)
            }
        }
        $this.excelConn.conn.Visible = $showWindow 
    }

    [void] SetWorksheet($num) {
        $this.excelConn.ws = $this.excelConn.workbook.Worksheets.Item($num)
    }

    [object] getWS() {
        $ws = if ($null -ne $this.excelConn.ws) {($this.excelConn.ws)} Else {($this.excelConn.workbook.Worksheets.Item(1))}
        return $ws;
    }

    [object] AddColumns([string] $cols) {
        $ws = $this.getWS()
        $ws.Columns("$cols").Insert()
        return $this
    }

    [boolean] compareFile($dir, $name) {
        $finalDir = $dir
        if ($finalDir -like "{*}") { 
            $finalDir = $this.meta[$dir.substring(1, $dir.Length -2)] 
        }
        $file2 = Get-Item "$finalDir\$name"
        return ( (Get-FileHash $this.file -Algorithm "SHA256").Hash -eq (Get-FileHash $file2 -Algorithm "SHA256").Hash )
    }
    [object] AddRows([string] $rows) {
        $ws = $this.getWS()
        $ws.Rows("$rows").Insert()
        return $this
    }

    [object] SetColumnWidth($cols, $size) {
        $ws = $this.getWS()
        $ws.Columns($cols).ColumnWidth = $size
        return $this
    }

    [object] SetCell($cells, $value) {
        $this.getWS().Range($cells).Value2 = $value
        return $this
    }

    [object] FillDown($cells) {
    
        $this.getWS().Range($cells).fillDown()
        return $this
    }

    [object] ColumnAutoFit($col) {
        return $this.getWS().Columns($col).AutoFit()
    }

        [string] GetCellValue($cell) {
        return $this.getWS().Range($cell).Value2
    }

    [object] AddFilter($row) {
        return $this.getWS().Rows($row).AutoFilter()
    }
    [System.Collections.DictionaryEntry] GetMeta() {
        return $this.meta
    }
    [void] SaveAndQuit() {
        $this.excelConn.workbook.Save()
        $this.excelConn.conn.Quit()
    }

    [int] GetRowCount() {
        return $this.getWS().UsedRange.Rows.Count
    }
    [void] SaveQuitAndMove($dir) {
        $this.SaveAndQuit()
        Start-Sleep -Seconds 1
        $newPath = $this.GetInProgressFile()
        Move-Item -Path $newPath -Destination $this.getMetaPath($dir)
    }

    [void] Quit() {
        $this.excelConn.conn.Quit()
        Start-Sleep -Milliseconds 500
    }
    [string] GetInProgressFile() {
        return "$($this.meta.inProgress)\$($this.file.BaseName).xlsx"
    }
    [boolean] NotExcelFile() {
        return ($null -eq $this.excelConn.conn)
    }
    [string] getMetaPath($string) {
        $finalDir = $string
        if ($finalDir -like "{*}") { 
            $finalDir = $this.meta[$string.substring(1, $string.Length -2)] 
        }
        return $finalDir
    }
    [void] MoveFileTo($dir)
    {
        Move-Item -Path $this.file -Destination $this.getMetaPath($dir)
    }

    [object] GetFile($dir, $name) {
        if ($this.DoesFileExist($dir, $name)) {
            return (Get-Item -Path "$($this.getMetaPath($dir))\$name")
        } else {
            Write-Error "File was not found."
            return $null
        }
    }
    [object] CreateFile($dir, $name, $txt) {
        Write-Host "path: $($this.getMetaPath($dir))"
        $LclFile = New-Item -Path "$($this.getMetaPath($dir))\$name"
        Add-Content $LclFile $txt
        return $LclFile
    }

    [object] DoesFileExist($dir, $filename) {
        return Test-Path -Path "$($this.getMetaPath($dir))\$filename"
    }
    [void] CopyFileTo($dir)
    {
        $finalDir = $this.getMetaPath($dir)
        Write-Host "Copying file to {completed}: $finalDir"
        Copy-Item -Path $this.file -Destination $finalDir
    }
    [void] ApplyFilter($byCol, $onRows, $filterName, [string]$fa1, [string]$fa2, [string]$fa3)  {
        $this.getWS().Rows($onRows).AutoFilter($byCol, @($filterName, $fa1, $fa2, $fa3), 7)
    }
}