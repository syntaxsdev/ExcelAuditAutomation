# Overview
This project utilizes two main files. AuditControl.ps1, AutomationMaker.ps1, and DataParser.ps1. This project aims to aid basic Excel automation, daily task automation for Excel, and the ability to move them to different folders when done processing.

# How To Create An Automation
1.	First, name your automation. Once you pick a name, open the AutomationMaker.ps1 file.
a.	Scroll down below the line that says automations go under here.
b.	Add the lines below, replacing “myAutomation” with your desire automation name.
```powershell
$excelScripts.newAutomation("MyAutomation", {
    $excel = $excelScripts.getExcel()
    
})
```
2.	Inside the function, use the documentation below to build the automation that suits your needs. If needed, 
$excelScripts.getWS() -> gets the Excel COM object for the default worksheet. Additional methods can be accessed through here.
3.	Make sure to SaveAndQuit() at the end of each automation function. If you are using multiple automations to recycle code or to achieve the baseline automation before adding more specific tasks, you do not have to add the SaveAndQuit() to that automation. You can recall a previously defined automation by `$excelScripts.run(“MyAutomation”)` at the top of your new automation’s code.
