Param(
[Parameter(Mandatory=$true)]$file,
[Parameter(Mandatory=$true)]$module
)

$pe = [PowerExcel]::new($file, $true)
. "$(Get-Location)\AutomationMaker.ps1" -action "run" -event $pe -module $module


class PowerExcel {
    [string] $file

    $excelConn = @{
        conn=$null
        workbook=$null
        ws=$null
    }

    PowerExcel($excelFile, $showWindow) {
        $this.file = $excelFile
        $this.excelConn.conn = (New-Object -ComObject Excel.Application)
        $this.excelConn.workbook = ($this.excelConn.conn.Workbooks.Open($excelFile))
        #save file as XLSX file version 51: xlOpenXMLWorkbook
        if ($excelFile.Extension -eq '.csv') {
            $this.excelConn.workbook.SaveAs($excelFile.FullName.Trim(".csv"), 51)

        }
        $this.excelConn.conn.Visible = $showWindow 
    }

    [void] SetWorksheet($num) {
        $this.excelConn.ws = $this.excelConn.workbook.Worksheets.Item($num)
    }

    [object] getWS() {
        $ws = if ($this.excelConn.ws -ne $null) {($this.excelConn.ws)} Else {($this.excelConn.workbook.Worksheets.Item(1))}
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
        $cellRange = $this.getWS().Range($cells).Value2 = $value
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
        $newPath = $this.file.Trim(".csv") + ".xlsx"
        Move-Item -Path $newPath -Destination $dir

    }
    [void] ApplyFilter($byCol, $onRows, $filterName, $filterAction)  {
        $this.getWS().Rows($onRows).AutoFilter($byCol, @($filterName, $filterAction), 7)
    }
}