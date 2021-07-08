function Start-Auto($days) {
    $config = Import-Clixml -Path "$(Get-Location)\savedConfig.xml"

    Write-Host "`nAuditingControl is running. There were $($config.count) automations found."
    $config.Keys | ForEach-Object {
        if ($config[$_].fileNames -ne "all" -or $config[$_].fileNames -ne "no") {
            $fileNames = $config[$_].fileNames -split ","
        }
        $filesInFolder = Get-ChildItem -Path $config[$_].original -Recurse -Include *.csv | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays(-$days)}
        $newMod = $_
        if ($newMod -like "*{ignore-path}*") {
            $newMod = $_.substring(0, $_.indexOf("{ignore-path}"))
        }
        foreach ($file in $filesInFolder) 
        {
            foreach ($name in $fileNames) {
                if ($file -like  "*$name*") {
                    Write-Host "File found matches criteria {$name}"
                    . "$(Get-Location)\DataParser.ps1" -file $file -module $newMod
                    #pause 2 seconds because some large datasets take longer to save and quit
                   if ($_ -eq $newMod) {
                    Start-Sleep -Seconds 1
                    Move-Item -Path "$($config[$_].original)\$($file.BaseName).xlsx" -Destination $config[$_].inProgress
                }
            }
           }
           
        }
    }
}

Start-Auto(Read-Host -Prompt "Enter the number of days to pull from (1 day is default)")
