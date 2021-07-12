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
        $this.meta = $import[($import.Keys | Where-Object {$_ -like "*$module*"})]

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
            $this.excelConn.workbook.SaveAs($newFile, 51)
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
        $finalDir = $dir
        if ($dir -like "{*}") { 
            $finalDir = $this.meta[$dir.substring(1, $dir.Length -2)] 
        }

        $this.SaveAndQuit()
        Start-Sleep -Seconds 1
        $newPath = $this.GetInProgressFile()
        Write-Host $newPath
        Move-Item -Path $newPath -Destination $finalDir
    }

    [string] GetInProgressFile() {
        return "$($this.meta.inProgress)\$($this.file.BaseName).xlsx"
    }
    [boolean] NotExcelFile() {
        return ($null -eq $this.excelConn.conn)
    }
    [void] CopyFileTo($dir)
    {
        $finalDir = $dir
        if ($dir -like "{*}") { 
            $finalDir = $this.meta[$dir.substring(1, $dir.Length -2)] 
        }
        Write-Host "Moving non excel file to {completed}: $finalDir"
        Copy-Item -Path $this.file -Destination $finalDir
    }
    [void] ApplyFilter($byCol, $onRows, $filterName, $filterAction)  {
        $this.getWS().Rows($onRows).AutoFilter($byCol, @($filterName, $filterAction), 7)
    }
}