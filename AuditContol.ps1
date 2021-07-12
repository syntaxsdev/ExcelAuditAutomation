function Start-Auto($date, $autoList) {
    $config = $autoList#Import-Clixml -Path "$(Get-Location)\savedConfig.xml"
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
            Write-Host "`nFile Found: [$file]"
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

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(750,210)
$Form.text                       = "Audit Control Task Runner"
$Form.TopMost                    = $false
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#9a8e9c")

$dateInputLbl                    = New-Object system.Windows.Forms.Label
$dateInputLbl.text               = "Enter the file creation date to run for the automation (ex: 07/09/2021):"
$dateInputLbl.AutoSize           = $true
$dateInputLbl.width              = 25
$dateInputLbl.height             = 10
$dateInputLbl.location           = New-Object System.Drawing.Point(7,9)
$dateInputLbl.Font               = New-Object System.Drawing.Font('Microsoft YaHei UI',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$dateInputLbl.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("")

$dateTxtBox                        = New-Object system.Windows.Forms.TextBox
$dateTxtBox.multiline              = $false
$dateTxtBox.Text                   = (Get-Date -format "MM/dd/yyyy")
$dateTxtBox.width                  = 128
$dateTxtBox.height                 = 20
$dateTxtBox.location               = New-Object System.Drawing.Point(500,7)
$dateTxtBox.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$startAuditTaskBtn               = New-Object system.Windows.Forms.Button
$startAuditTaskBtn.text          = "Start Automations"
$startAuditTaskBtn.width         = 135
$startAuditTaskBtn.height        = 30
$startAuditTaskBtn.location      = New-Object System.Drawing.Point(255,166)
$startAuditTaskBtn.Font          = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$startAuditTaskBtn.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#bbb5c6")

$forceCloseBtn               = New-Object system.Windows.Forms.Button
$forceCloseBtn.text          = "Force Close Excel Connections"
$forceCloseBtn.width         = 225
$forceCloseBtn.height        = 30
$forceCloseBtn.location      = New-Object System.Drawing.Point(10,166)
$forceCloseBtn.Font          = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$forceCloseBtn.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#bbb5c6")

$Form.controls.AddRange(@($dateInputLbl,$dateTxtBox,$startAuditTaskBtn, $forceCloseBtn))

$forceCloseBtn.Add_Click({Stop-Process -Name 'excel'})
$startAuditTaskBtn.Add_Click({ startAutomation })

function startAutomation() {
    $autoList = $automationsChosen
    $checkBoxes.Keys | ForEach-Object {
        $chkBox = $checkBoxes[$_]
        if ($false -eq $chkBox.Checked) {
            $autoList.Remove($_)
        } else {
            $autoList[$_] = $automationsChosen[$_]
    }
    }
    $date =  [Datetime]::ParseExact($dateTxtBox.Text, 'MM/dd/yyyy', $null)
    Start-Auto -date $date -autoList $autoList
    $autoList = $automationsChosen
}

$config = Import-Clixml -Path "$(Get-Location)\savedConfig.xml"
$automationsChosen = $config
$checkBoxes = @{}
$xPos = 0
$yPos = 30
$config.Keys | ForEach-Object {
    $TempChkBx                   = New-Object system.Windows.Forms.CheckBox
    $TempChkBx.text              = $_
    $TempChkBx.AutoSize          = $false
    $TempChkBx.width             = 250
    $TempChkBx.height            = 20
    $TempChkBx.Checked           = $true
    $TempChkBx.location          = New-Object System.Drawing.Point($xPos, $yPos)
    $TempChkBx.Font              = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
    $TempChkBx.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
    #$automationsChosen[$_] = $config[$_]
    $checkBoxes[$_] = $tempChkBx

    $Form.controls.Add($TempChkBx)
    
    $xPos += 250
    if ($xPos -gt 600) {
        $xPos = 0
        $yPos += 30
    }
}

#show the dialog form
$Form.ShowDialog()