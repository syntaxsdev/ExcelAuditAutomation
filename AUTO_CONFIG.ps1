#AUTO_CONFIG
$automations = @{}

while ( $autoName = (Read-Host -Prompt "Please enter the excel Job function name you are staging. If there are no more, type 'none'")) {
    if ($autoName -eq 'none') { Write-Host 'Successfully added all automations.'; break}

    $original = Read-Host -Prompt "Enter ORIGINAL path for automation- $autoName"
    $inProgress = Read-Host -Prompt "Enter IN PROGRESS path for automation- $autoName"
    $completed = Read-Host -Prompt "Enter COMPLETED path for automation- $autoName"

    $automations[$autoName] = @{original=$original 
                                inProgress=$inProgress
                                completed=$completed}
    Write-Host ("  ")
}

$automations | Export-Clixml -Path ("$(Get-Location)\savedConfig.xml") 