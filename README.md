# ExcelAuditAutomation Documentation
Video Demonstration [HERE!](https://www.youtube.com/watch?v=xo9TvaocDtg&ab_channel=Syntaxs)
# Overview
This project utilizes three main files. **AuditControl.ps1**, **AutomationMaker.ps1**, and **DataParser.ps1**. This project aims to aid basic Excel automation for daily tasks and auditing, with the ability to stage them to different folders when done processing.


# What does this do?
### Scenario
Every day you have a background process that drops two excel files into a network folder you must process for audit. You add rows, additional columns, do sorting, filtering, and pretty much the same task the same way everyday. You could use a macro, but say you have multiple files, across multiple folders that you do daily automations for, that can become hard to manage.

With this ExcelAuditAutomation program, you can predesign a workflow to do all of those tasks, run them on specific file names and file types, and even work with non-excel files. 

## Is this solution right for your process?
### If you can answer yes to these question, this solution will work for you.
*Your file is in a .CSV or Excel format*

*The files in the folders you are working on include moving/comparing/copying additional non excel files*

*You use the same or similar formulas based on data inside the excel file or non excel file*

*You would like to reduce time and error spent doing frequent manual processes on excel files*


# Things to Note
- .CSV and .XLSX files do not need additional parameters. In your configuration, you can set what type of files to run the automation on. If one day someone converts the .csv to an .xlsx, the program will work the same without any additional requirements. You may even specify whether or not you would like the file to convert to .xlsx, or stay in .csv format. NOTE that this should only be used for completing excel automations and staying open for audit purposes, but saving a non-converted excel file will not save anything done inside excel.
- You can process non excel files in a folder in a seperate automation, or its own automation. More details on this will be explained in #Important Notes

# Instructions for setup
1. Create your automations first. Skip to the step below on how to create one.

2.	The next step is to run the configuration maker which is the **CONFIG GUI.ps1** file. 
    - Enter each automation function name, and then the three associated folders *(ORIGINAL, IN PROGRESS, COMPLETE)*
    - ORIGINAL and IN PROGRESS paths CANNOT be blank, however COMPLETE path can be empty.
    - To whitelist a file so only certain file names will be processed, enter a keywords, seperated by a comma.
    - If you do not want your file to convert CSV to XLSX, and want it to open out the ORIGINAL folder, untick "Convert to Excel". This will not convert any other files besides .CSVs to .XLSX.
    - When done with your config, remember to hit **ADD MODULE** first. It will add that config module and reset the fields to blank. You can add as many modules as needed. When done, hit **SAVE CONFIG**. This will generate an XML to same folder the GUI is in.
        - NOTE: It may be easier to edit the XML file to change one specific path, than to import the existing config and override.
        - If you want to add more to the existing config, hit the "Import" button. If one does not already exist it will give you a warning. Importing and adding the same automation name will override the existing config for that automation.
       
3.	Once **CONFIG GUI.ps1** is done, open **AuditControl.ps1** to start your automation. With the attached configuration saved, all you have to do now is type the file creation date that the audit files were made on. At the time, you can only process one day at a time. 
4. Hit Start. It will now process the automation you specified on the folders you configured, with the specified names and specified file types.

# How To Create An Automation
1.	First, name your automation. Once you pick a name, edit the **AutomationMaker.ps1** file.
    - Scroll down below the line that says *----AUTOMATIONS GO UNDER HERE----*
    - Add the lines below, replacing “myAutomation” with your desire automation name
```powershell
$excelScripts.newAutomation("MyAutomation", {
    $excel = $excelScripts.getExcel()
    #AUTOMATION GOES HERE
})
```
2.	Inside the function, use the documentation below to build the automation that suits your needs. If needed, 
$excelScripts.getWS() -> gets the Excel COM object for the default worksheet. Additional methods can be accessed through here.
3.	Make sure to use the commands `SaveAndQuit(), Quit(), or SaveQuitAndMove("{folder}")` at the end of each automation function if you would like the function to exit/save when done, otherwise it will stay open. If you are using multiple automations to recycle code or to achieve the same baseline automation before adding more specific tasks, you do not have to add the save or quit methods to that automation. You can recall a previously defined automation by `$excelScripts.run(“MyAutomation”)` at the top of your new automation’s code.

### Example 1 of a basic program
This automation will create a new Column D, and sum columns B and C and fill down.
```powershell
$excelScripts.newAutomation("MyAutomation", {
    $excel = $excelScripts.getExcel()
    $excel.AddColumns("D:D")
    $excel.SetCell("D1", "Sum")
    $excel.SetCell("D1", ($excel.GetCellValue("B1") + $excel.GetCellValue("C1")))
    $excel.FillDown("D1:D$($excel.GetRowCount())")
    $excel.SaveQuitAndMove("{completed}")
})
```
Lets break down this code.

```powershell 
$excelScripts.newAutomation("MyAutomation", {
```

This line creates a new automation called "myAutomation" with starting brackets to enclose automation code inside.

```powershell 
$excel = $excelScripts.getExcel() 
```
Grabs the excel automation object. Essential to do anything with the file.

```powershell 
$excel.AddColumns("D:D")
```

Add a columns to Column D only. If you want to add a column to D-F, use "D:F"

```powershell 
$excel.SetCell("D1", "Sum")
```

Sets cell's text D1 to "Sum". The function can also be used to set formulas.

```powershell 
$excel.SetCell("D2", ($excel.GetCellValue("B2") + $excel.GetCellValue("C2")))
```

Sets the starting cell D2 to be the sum of B2 and C2

```powershell 
$excel.FillDown("D2:D$($excel.GetRowCount())")
```

Autofills the formula from D2 to all the available rows.

```powershell 
$excel.SaveQuitAndMove("{completed}")
```

Saves the file, exits, and moves it to the {completed} folder that is specified in the configuration. You can also use any path you want, it does not have to save to the configuration paths.

### Example 2 of a basic program
This automation will take all files in the folder including .txt, .log, .csv/or excel files and process them for automation.
```powershell
$excelScripts.newAutomation("SQLJobsAudit", {
    $excel = $excelScripts.getExcel()
    if ($excel.NotExcelFile()) {
        $excel.CopyFileTo("{completed}")
        return
    }
    
    $excel.SetCell("A1", "Ticket Nbr")
    $excel.SetCell("B1", "Date")
    $excel.SetCell("C1", "Time")

    $excel.SaveQuitAndMove("{inProgress}")
})
```
Lets break down this code.

```powershell
if ($excel.NotExcelFile()) {
        $excel.CopyFileTo("{completed}")
        #return
    }
```
In this code block, the file will check to see if the current file is NOT an excel file, since we specified in the configuration settings we want to process all the files, including non excel files. 
Once it confirms that the current file is not an excel file (either .csv or .xlsx) it will move that file to the {completed} path which was determined in the configuration folder. It will then return, ending it's automation there.

```powershell
    $excel.SetCell("A1", "Ticket Nbr")
    $excel.SetCell("B1", "Date")
    $excel.SetCell("C1", "Time")

    $excel.SaveQuitAndMove("{inProgress}")
```
This code block will do some basic functions, such as setting the cells A-C1. It will then move all the excel files into the {inProgress} folders.

### Advanced Usage Example
This is an advanced usage of the automation. The original folder for this file has multiple .csv/.xlsx files and multiple .sql files for auditing purposes.
We want to quit the file as soon as it opens, only for grabbing the name of the audit file. We then compare that the .CSV file to the .CSV generated yesterday.
```powershell
$excelScripts.newAutomation("SqlDailyVerification", {
    $excel = $excelScripts.getExcel()
    $fileName = $excel.file.Name
    $fileExt = $excel.file.Extension
    $excel.Quit()
    $currentDate = $fileName.substring($fileName.LastIndexOf('_')+1, 8)
    $yesterday = (Get-Date ([Datetime]::ParseExact($currentDate, 'yyyyMMdd', $null).AddDays(-1))).ToString("yyyyMMdd")
    $newFile = "$($fileName.substring(0, $fileName.LastIndexOf("_")+1))$yesterday$fileExt"
    $compareResult = $excel.CompareFile("{original}", $newFile)
    $logFileName = "SqlDaily_CompareLog_$currentDate.csv"

    if ($false -eq ($excel.DoesFileExist("{inProgress}",$logFileName))) {
        $lclFile = $excel.CreateFile("{inProgress}", $logFileName, "file,status,message")
    } else { $lclFile = $excel.GetFile("{inProgress}", $logFileName) }

    if ($true -eq $compareResult) {
        Add-Content $lclFile "$fileName, SAME, [$fileName] is the same as [$newFile]"
    } elseif ($false -eq $compareResult) {
        Add-Content $lclFile "$fileName, CHANGED, [$fileName] IS NOT THE SAME AS [$newFile] OR DOES NOT EXIST!!"
    }
    $excel.CopyFileTo("{completed}")
})
```
Lets break down this advanced code.
```powershell
 $excel = $excelScripts.getExcel()
    $fileName = $excel.file.Name
    $fileExt = $excel.file.Extension
    $excel.Quit()
```
We don't actually want to touch the file, we just want the name of it. However - at this time, by default all processable files (.csv or .xlsx) will open, so we close it by using Quit() - which does not commit changes.
```powershell
    $currentDate = $fileName.substring($fileName.LastIndexOf('_')+1, 8)
    $yesterday = (Get-Date ([Datetime]::ParseExact($currentDate, 'yyyyMMdd', $null).AddDays(-1))).ToString("yyyyMMdd")
    $newFile = "$($fileName.substring(0, $fileName.LastIndexOf("_")+1))$yesterday$fileExt"
    $compareResult = $excel.CompareFile("{original}", $newFile)
```
We then get date of the file, which is in the file name formatted as "SQLDailyChanges_20210722.CSV" for example. We know this ahead of time, and we grab this name, subtract one using the DateTime object. We then use that time to grab the yesterday file. Note that `.CompareFile()` will work on any file going through the automation, however the file cannot be open.
We compare the current file to the yesterday file of the same format that is located in the {original} folder. 

```powershell
    $logFileName = "SqlDaily_CompareLog_$currentDate.csv"

    if ($false -eq ($excel.DoesFileExist("{inProgress}",$logFileName))) {
        $lclFile = $excel.CreateFile("{inProgress}", $logFileName, "file,status,message")
    } else { $lclFile = $excel.GetFile("{inProgress}", $logFileName) }

    if ($true -eq $compareResult) {
        Add-Content $lclFile "$fileName, SAME, [$fileName] is the same as [$newFile]"
    } elseif ($false -eq $compareResult) {
        Add-Content $lclFile "$fileName, CHANGED, [$fileName] IS NOT THE SAME AS [$newFile] OR DOES NOT EXIST!!"
    }
    $excel.CopyFileTo("{completed}")
```
Now we will create a .CSV log file for today's entries. However since the automation is looped, we will only create the file if one for today hasn't been made yet, so the first automation to run this for the day will create the file. We will store the file to the variable `$lclFile` and if the file already exists, we will use the `.GetFile()` to grab that already existing .CSV file.
Next, we append the .CSV file based on the result from the `.CompareFile()` method. Once we are done appending, we will then use `.CopyFileTo()`to create a copy of this file {completed} folder, as we would like to save a copy in the original as well.
#### Note once you quit the excel instance, you can only use the `.MoveFileTo()` or `.CopyFileTo()` methods to change that file's location.

#### A list of additional available commands below.

## Important Notes
- When adding an automation, you must run the **CONFIG GUI.ps1** to generate a new configuration for your automation and attach the paths. 

## Functionality
```powershell
+ SetWorksheet($num)
+ getWS()
+ AddColumns($cols)
+ AddRows($rows)
+ SetColumnWidth($cols, $width)
+ SetCell($cells, $value)
+ GetCellValue($cell)
+ FillDown($cells)
+ ColumnAutoFit($cols)
+ AddFilter($rows)
+ ApplyFilter($byCol, $onRows, $filterName, $filterAction1, $filterAction2, $filterAction3)
+ GetRowCount()
+ SaveAndQuit()
+ Quit()
+ SaveQuitAndMove($dir)

#Non Excel file functions
+ NotExcelFile()
+ compareFile($configDir, $name)
+ CopyFileTo($configDir)
+ MoveFileTo($configDir, $name)
+ DoesFileExist($configDir, $name)
+ GetFile($configDir, $name)
+ CreateFile($configDir, $name, $data)
+ GetFile($config)
```
## More Examples

```powershell
$excel.AddColumns("A:C") # Adds columns to in spot A-C. shifting over everything to right

$excel.SetColumnWidth("A:A", 19) # Sets Column A only to size 19.
$excel.FillDown("A1:A20") # Autofills from the first row down.
$excel.ApplyFilter(2, "1:1", "Task Type", "Backup") # 2 corresponds to the column that has the filter, "1:1" corresponds to which row
#the entire filter grouping is on, "Task Type" is the Filter field name, and "Backup" is what I want to filter by.
 ```
 
