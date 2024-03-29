﻿<#
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
   $excel.SetCell("C1", "Sum")
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
#---------------------------------------------AUTOMATIONS GO UNDER HERE---------------------------------------------------------

$excelScripts.newAutomation("SqlJobsCancelled", {
    
})




#--------------------------------------------------AUTOMATIONS END HERE--------------------------------------------------------
switch ($action) {
    "run" {try {
        Write-Host "Attempting to run Module [$module]"
        $excelScripts.run($module)
        Write-Host "Successfully executed Module [$module]!!`n"
        . "$(Get-Location)\logger.ps1" -log "Successfully executed Module [$module]"
        } catch {
            . "$(Get-Location)\logger.ps1" -log "An error occured in attempting to run the automation for Module [$module]"}}

}


class ExcelAutomation {
    $automations = @{}
    $excelObj = $null
    $meta = @{}

    ExcelAutomation($excel) {
        $this.excelObj = $excel
    }

    [object] getExcel() {
        return $this.excelObj
    }

    [void] newAutomation($name, [scriptblock] $func) {
        $this.automations[$name] = $func
        $this.automations[$name]
    }
    
    [void] run($name) {
        $this.automations[$name].Invoke($this.getExcel())
        #$this.getExcel().SaveAndQuit()
    }
}
