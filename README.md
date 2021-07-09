# Overview
This project utilizes three main files. **AuditControl.ps1**, **AutomationMaker.ps1**, and **DataParser.ps1**. This project aims to aid basic Excel automation for daily tasks and auditing, with the ability to stage them to different folders when done processing.

# Instructions for setup
1. Create your automations first. Skip to the step below on how to create one.

2.	The next step is to run the configuration maker which is the **GUI.ps1** file. CLI based configuration is the **AUTO_CONFIG.ps1**.
    - Enter each automation function name, and then the three associated folders *(ORIGINAL, IN PROGRESS, COMPLETE)*
    - ORIGINAL and IN PROGRESS paths CANNOT be blank, however COMPLETE path can be empty.
    - To whitelist a file so only certain file names will be processed, enter a keywords, seperated by a comma.
    - If you want to add more to the existing config, hit the "Import" button. If one does not already exist it will give you a warning.
    
3.	Once **AUTO_CONFIG.ps1** is done, open **AuditControl.ps1** to start your automation. With the attached configuration saved, all you have to do now is type the amount of days back from the current day you want the program to process the files. 
EX: If it is a Monday, you must type 3, to process audits from Friday, Saturday & Sunday.
If it is any other day of the week, you may only need to type 1 for one day.


# How To Create An Automation
1.	First, name your automation. Once you pick a name, open the **AutomationMaker.ps1** file.
    - Scroll down below the line that says automations go under here.
    - Add the lines below, replacing “myAutomation” with your desire automation name.
```powershell
$excelScripts.newAutomation("MyAutomation", {
    $excel = $excelScripts.getExcel()
    #AUTOMATION GOES HERE
})
```
2.	Inside the function, use the documentation below to build the automation that suits your needs. If needed, 
$excelScripts.getWS() -> gets the Excel COM object for the default worksheet. Additional methods can be accessed through here.
3.	Make sure to SaveAndQuit() at the end of each automation function. If you are using multiple automations to recycle code or to achieve the same baseline automation before adding more specific tasks, you do not have to add the SaveAndQuit() to that automation. You can recall a previously defined automation by `$excelScripts.run(“MyAutomation”)` at the top of your new automation’s code.
4.	
### Example of a basic program
This automation will create a new Column D, and sum columns B and C, filled down.
```powershell
$excelScripts.newAutomation("MyAutomation", {
    $excel = $excelScripts.getExcel()
    $excel.AddColumns("D:D")
    $excel.SetCell("D1", ($excel.GetCellValue("B1") + $excel.GetCellValue("C1")))
    $excel.FillDown("D1:D$($excel.GetRowCount())")
    $excel.SaveQuitAndMove("{completed}")
})
```


## Notes
- When adding an automation, you must run the **AUTO_CONFIG.ps1** to generate a new configuration for your automation and attach the paths. 
- When using SaveAndQuit(), it will automatically move the file to the IN PROGRESS folder. If you want to disable this, when generating the config in **AUTO_CONFIG.ps1** use _"myAutomation{ignore-path}"_ as the automation name, replacing "myAutomation" with the name of your function.
- If you are using _SaveQuitAndMove($dir)_, you must add "{ignore-path}" so the code knows you will not be changing its initial determined route.
- If you want to save the file to a different location, use the function SaveQuitAndMove($dir). If you want to change it to the completed folder, use the `("{completed}")` tag or if you want to keep it in the ORIGINAL folder, then you must still use _"{ignore-path}"_ and use the method `SaveAndQuit()`

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
$excel.ApplyFilter(2, "1:1", "Task Type", "Backup") # 2 corresponds to the column that has the filter, "1:1" corresponds to which row
#the entire filter grouping is on, "Task Type" is the Filter field name, and "Backup" is what I want to filter by.
 ```
 
