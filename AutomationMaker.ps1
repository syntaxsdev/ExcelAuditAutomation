<#
THIS IS WHERE ALL EXCEL AUTOMATION SHOULD BE PLACED
EXAMPLES GIVEN BELOW
#>

Param([Parameter(Mandatory=$true)] [string]$action,
       [Parameter(Mandatory=$true)] [object] $event,
       [Parameter(Mandatory=$true)] [string] $module)

$excelScripts = [ExcelAutomation]::new($event)

<#
    Write functions below if you are not going use anonymous functions.
#>

function SqlJobsCancelled($excel) {
   
}

function Netwrix($excel) {

}


<#
    Add new automations below, following the existing formatting
    newAutomation( NAME, $function:NAME OF FUNCTION ) 

    You can add a function TWO ways:
    1.anonymously, and then get the $excel object using $excelScripts.getExcel() or ;
    2. using the function name to in place and passing $excel as the parameter
#>
$excelScripts.newAutomation("SqlJobsCancelled", {
    $excel = $excelScripts.getExcel()

    $excel.SetCell("A1", "Ticket Nbr")
    $excel.SetCell("B1", "Date")
    $excel.SetCell("C1", "Time")
    $excel.SetCell("D1", "Job Name")
    $excel.SetCell("E1", "Job Step")
    $excel.SetCell("F1", "Message")
    $excel.ColumnAutoFit("A:E")

    if ($excel.GetCellValue("D2") -eq '') {
        $excel.SetCell("A2", "No Data")
        return
    }
    $excel.AddFilter("1:1")
})



$excelScripts.newAutomation("Netwrix", {
    $excel = $excelScripts.getExcel()

    $excel.AddColumns("A:B")
    $excel.SetCell("A1", "Ticket")
    $excel.SetCell("B1", "Special Notes")
    $excel.SetColumnWidth("A:A", 15)
    $excel.ColumnAutoFit("B:C")
    $excel.ColumnAutoFit("F:M")
    $excel.ColumnAutoFit("P:P")

    $excel.AddFilter("1:1")
    If ($excel.GetCellValue("C2") -eq '') {
        $excel.SetCell("A2", "No Data")
        return
    }

})

$excelScripts.newAutomation("Autofill", {
    $excelScripts.run("Netwrix") #recycle method
    $excel = $excelScripts.getExcel()
    
    If ($excel.GetCellValue("C3") -ne '') {
        $numRows = $excel.GetRowCount()
        $excel.SetCell("A2", '=IFNA(IFS(I2="CVEN\Nebula_svc","Pipeline",I2="CVEN\devopsprod","Pipeline",ISNUMBER(SEARCH("GO_",G2))=TRUE,"GreenShades",AND(ISNUMBER(SEARCH("dbo.GS",G2))=TRUE,RIGHT(I2,2)="GS"),"GreenShades",AND(COUNTIFS(J:J,J2,I:I,I2)>1,RIGHT(I2,2)="GS"),"GreenShades"),"")')
        $excel.FillDown("A2:A$numRows")
        $excel.SetCell("B2", '=IFNA(IFS(I2="CVEN\devopsprod","DevOps Pipeline Deployment",A2="GreenShades","GreenShades Process",I2="CVEN\Nebula_svc","DevOps Pipeline Deployment – Service Broker"),"")')
        $excel.FillDown("B2:B$numRows")
        $excel.SaveAndQuit()
    }


})

$excelScripts.newAutomation("Backup", {
    $excel = $excelScripts.getExcel()
    $excel.AddRows("1:23")
    $excel.SetCell("A1", "SLA Domain")
    $excel.SetCell("A2", "Daily_CTG-UCSQLAAG01")
    $excel.SetCell("A3", "Daily_CTG-UCSQLAAG02")
    $excel.SetCell("A4", "Daily_CTG-SQL05")
    $excel.SetCell("A5", "Daily_CTG-SQL08")
    $excel.SetCell("A6", "Daily_CTG-SQLGPCL01")
    $excel.SetCell("A7", "Daily_CTG-DBA01")
    $excel.SetCell("A8", "Daily_CTG-SQLAAG01")
    $excel.SetCell("A9", "Daily_CTG-SQLAAG02")
    $excel.SetCell("A10",  "DEV-SQL01")
    $excel.SetCell("A11",  "DEV-SQL02")
    $excel.SetCell("A12",  "DEV-SQLGP01")
    $excel.SetCell("A13",  "DEV-SQLGP02")
    $excel.SetCell("A14",  "QA-SQL01")
    $excel.SetCell("A15",  "QA-SQL02")
    $excel.SetCell("A16",  "QA-SQLGP01")
    $excel.SetCell("A17",  "Daily_DR-SQL01")
    $excel.SetCell("A18",  "Daily_DR-SQL02")
    $excel.SetCell("A19",  "Daily_DR-SQLGP01")
    $excel.SetCell("A20",  "DEV-SQL-TM1")
    $excel.SetCell("A21",  "QA-SQL-TM1")
    $excel.SetCell("A22",  "Daily_HQ-SQL-TM1")
    $excel.SetCell("A23",  "Unprotected")

    $excel.ColumnAutoFit("A:A")
    $excel.SetCell("B2", "=COUNTIFS(C,""BACKUP"",C[-1],""<>Succeeded"",C[4],RC[-1])")
    $excel.FillDown("B2:B23")
    $excel.AddFilter("24:24")
    $excel.ApplyFilter(2, "24:24", "Task Type", "Backup")
    $excel.ApplyFilter(1, "24:24", "Task Status", "Failed")
    $excel.SaveQuitAndMove("C:\Users\glisid\Dev")
})


$excelScripts.newAutomation("SqlDaily", {
    $excel = $excelScripts.getExcel()

    
    $excel.AddColumns("A:A")
    
    $excel.SetCell("A1", "Ticket Number").
        SetCell("B1", "Server").
        SetCell("C1", "EventTime").
        SetCell("D1", "Database").
        SetCell("E1", "Schema").
        SetCell("F1", "Object").
        SetCell("G1", "Action").
        SetCell("H1", "Action_Status").
        SetCell("I1", "UserName").
        SetCell("J1", "Statement").
        SetCell("K1", "Additional_Info")


        $excel.ColumnAutoFit("A:B")
        $excel.SetColumnWidth("C:C", 19)
        $excel.SetColumnWidth("F:F", 19)
        
        If ($excel.GetCellValue("E3") -eq '' -and $excel.GetCellValue("E4") -eq '') {
            $excel.SetCell("A2", "No Data")
        }
        if ($excel.GetCellValue("E2") -eq '') {
            $excel.FillDown("A2:A4")
        }
        $excel.AddFilter("1:1")
})



















switch ($action) {

    "run" {try {
        Write-Host "Attempting to run Module [$module]"
        $excelScripts.run($module)
        . "$(Get-Location)\logger.ps1" -log "Successfully executed Module [$module]"
        } catch {
            . "$(Get-Location)\logger.ps1" -log "An error occured in attempting to run the automation for Module [$module]"}}


}


class ExcelAutomation {
    $automations = @{}
    $excelObj = $null

    ExcelAutomation($excel) {
        $this.excelObj = $excel
    }

    [object] getExcel() {
        return $this.excelObj
    }

    [object] newAutomation($name, [scriptblock] $func) {
        $this.automations[$name] = $func
        return $this.automations[$name]
    }
    
    [void] run($name) {
        $this.automations[$name].Invoke($this.getExcel())
        #$this.getExcel().SaveAndQuit()
    }
}