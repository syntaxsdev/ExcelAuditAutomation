Param(
$log,
$moduleName
)
Add-Content "$(Get-Location)\log.txt" "$(Get-Date) - MESSAGE: $log"