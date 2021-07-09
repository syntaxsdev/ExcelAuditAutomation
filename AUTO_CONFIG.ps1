#AUTO_CONFIG
$automations = @{}

while ( $autoName = (Read-Host -Prompt "Please enter the excel Job function name you are staging. If there are no more, type 'none'")) {
    if ($autoName -eq 'none') { Write-Host 'Successfully added all automations.'; break}

    $original = Read-Host -Prompt "Enter ORIGINAL path for automation - $autoName"
    $inProgress = Read-Host -Prompt "Enter IN PROGRESS path for automation - $autoName"
    $completed = Read-Host -Prompt "Enter COMPLETED path for automation - $autoName"
    $specific = Read-Host -Prompt "Are there some files names it should only run on? Enter keywords of file names followed by a comma if there are multiple, if there are none, enter `"all`""

    $automations[$autoName] = @{original=$original 
                                inProgress=$inProgress
                                completed=$completed
                                fileNames=$specific
                            }
}

$automations | Export-Clixml -Path ("$(Get-Location)\savedConfig.xml") 