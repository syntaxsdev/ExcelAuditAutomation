function Start-Auto($date) {
    $config = Import-Clixml -Path "$(Get-Location)\savedConfig.xml"

    Write-Host "`nAuditingControl is running. There were $($config.count) automations found."
    $config.Keys | ForEach-Object {
        if ($config[$_].fileNames -ne "all" -or $config[$_].fileNames -ne "no" -or $config[$_].fileNames -ne "") {
            $fileNames = $config[$_].fileNames -split ","
        }
        $filesInFolder = Get-ChildItem -Path $config[$_].original -Recurse -Include ($config[$_].fileType) | Where-Object {$_.CreationTime -ge $date -and $_.CreationTime -le $date.AddDays(1)}
        $newMod = $_
        if ($newMod -like "*{ignore-path}*") {
            $newMod = $_.substring(0, $_.indexOf("{ignore-path}"))
        }
        foreach ($file in $filesInFolder) 
        {
            foreach ($name in $fileNames) {
                $trimName = $name.Trim()
                if ($file -like  "*$trimName*") {
                    #explains whats happening
                    if ($trimName -eq "") { Write-Host "File found has no name criteria using file extension {$($config[$_].fileType)}"
                    } else { Write-Host "File found matches criteria {$trimName} using file extension {$($config[$_].fileType)}" }
                    
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

Start-Auto([Datetime]::ParseExact( (Read-Host -Prompt "Enter the file creation date (ex: 07/09/2021)"), 'MM/dd/yyyy', $null))