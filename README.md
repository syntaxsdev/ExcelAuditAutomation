# Overview
This project utilizes two main files. AuditControl.ps1, AutomationMaker.ps1, and DataParser.ps1. This project aims to aid basic Excel automation, daily task automation for Excel, and the ability to move them to different folders when done processing.

# How To Create An Automation
1.	First, name your automation. Once you pick a name, open the AutomationMaker.ps1 file.
a.	Scroll down below the line that says automations go under here.
b.	Add the lines below, replacing “myAutomation” with your desire automation name.
```powershell
$excelScripts.newAutomation("MyAutomation", {
    $excel = $excelScripts.getExcel()
    #AUTOMATION GOES HERE
})
```
2.	Inside the function, use the documentation below to build the automation that suits your needs. If needed, 
$excelScripts.getWS() -> gets the Excel COM object for the default worksheet. Additional methods can be accessed through here.
3.	Make sure to SaveAndQuit() at the end of each automation function. If you are using multiple automations to recycle code or to achieve the baseline automation before adding more specific tasks, you do not have to add the SaveAndQuit() to that automation. You can recall a previously defined automation by `$excelScripts.run(“MyAutomation”)` at the top of your new automation’s code.


## Functionality
```powershell
+ SetWorksheet($num)
+ AddColumns($cols)
+ AddRows($rows)
+ SetColumnWidth($cols, $width)
+ SetCell($cells, $value)
+ GetCellValue($cell)
+ FillDown($cells)
+ ColumnAutoFit($cols)
+ AddFilter($rows)
+ ApplyFilter($byCol, $onRows, $filterName, $filterAction)
+ GetRowCount()
+ SaveAndQuit()
+ SaveQuitAndMove($dir)
```
## Examples

```powershell
$excel.AddColumns("A:C") # Adds columns to in spot A-C. shifting over everything to right

$excel.SetColumnWidth("A:A", 19) # Sets Column A only to size 19. 
$excel.FillDown("A1:A20") # Autofills from the first row down.
$excel.ApplyFilter(2, "1:1", "Task Type", "Backup") # 2 corresponds to the column that has the filter, "1:1" corresponds to which row the entire filter grouping is on, "Task Type" is the Filter field name, and "Backup" is what I want to filter by.
 ```
 
