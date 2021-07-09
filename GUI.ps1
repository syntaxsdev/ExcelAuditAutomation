<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    bak
.SYNOPSIS
    Form
#>
$global:config = @{}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(711,295)
$Form.text                       = "Excel Automation Config"
$Form.TopMost                    = $false
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#8b7c7c")

$jobTxtBox                       = New-Object system.Windows.Forms.TextBox
$jobTxtBox.multiline             = $false
$jobTxtBox.width                 = 282
$jobTxtBox.height                = 20
$jobTxtBox.location              = New-Object System.Drawing.Point(401,12)
$jobTxtBox.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$jobLbl                          = New-Object system.Windows.Forms.Label
$jobLbl.text                     = "Please enter the excel Job function name you are staging:"
$jobLbl.AutoSize                 = $true
$jobLbl.width                    = 25
$jobLbl.height                   = 10
$jobLbl.location                 = New-Object System.Drawing.Point(12,15)
$jobLbl.Font                     = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$jobLbl.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$jobLbl.Tool

$originalLbl                     = New-Object system.Windows.Forms.Label
$originalLbl.text                = "Enter ORIGINAL path for automation:"
$originalLbl.AutoSize            = $true
$originalLbl.width               = 25
$originalLbl.height              = 10
$originalLbl.location            = New-Object System.Drawing.Point(12,47)
$originalLbl.Font                = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$originalLbl.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$originalTxtBox                  = New-Object system.Windows.Forms.TextBox
$originalTxtBox.multiline        = $false
$originalTxtBox.width            = 397
$originalTxtBox.height           = 20
$originalTxtBox.location         = New-Object System.Drawing.Point(286,44)
$originalTxtBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$AddModuleBtn                    = New-Object system.Windows.Forms.Button
$AddModuleBtn.text               = "Add Module"
$AddModuleBtn.width              = 120
$AddModuleBtn.height             = 30
$AddModuleBtn.location           = New-Object System.Drawing.Point(20,250)
$AddModuleBtn.Font               = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$AddModuleBtn.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$ImportBtn                       = New-Object system.Windows.Forms.Button
$ImportBtn.text                  = "Import Existing Config"
$ImportBtn.width                 = 180
$ImportBtn.height                = 30
$ImportBtn.location              = New-Object System.Drawing.Point(280,250)
$ImportBtn.Font                  = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$ImportBtn.ForeColor             = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$completeLbl                     = New-Object system.Windows.Forms.Label
$completeLbl.text                = "Enter COMPLETE path for automation:"
$completeLbl.AutoSize            = $true
$completeLbl.width               = 25
$completeLbl.height              = 10
$completeLbl.location            = New-Object System.Drawing.Point(13,122)
$completeLbl.Font                = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$completeLbl.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$completeTxtBox                  = New-Object system.Windows.Forms.TextBox
$completeTxtBox.multiline        = $false
$completeTxtBox.width            = 397
$completeTxtBox.height           = 20
$completeTxtBox.location         = New-Object System.Drawing.Point(286,118)
$completeTxtBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$inProgressLbl                   = New-Object system.Windows.Forms.Label
$inProgressLbl.text              = "Enter IN PROGRESS path for automation:"
$inProgressLbl.AutoSize          = $true
$inProgressLbl.width             = 25
$inProgressLbl.height            = 10
$inProgressLbl.location          = New-Object System.Drawing.Point(13,83)
$inProgressLbl.Font              = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$inProgressLbl.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$inProgressTxtBox                = New-Object system.Windows.Forms.TextBox
$inProgressTxtBox.multiline      = $false
$inProgressTxtBox.width          = 397
$inProgressTxtBox.height         = 20
$inProgressTxtBox.location       = New-Object System.Drawing.Point(286,80)
$inProgressTxtBox.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$keywordFileLbl                  = New-Object system.Windows.Forms.Label
$keywordFileLbl.text             = "Enter keywords for the file names you only want the program to run on, separated by commas."
$keywordFileLbl.AutoSize         = $true
$keywordFileLbl.width            = 25
$keywordFileLbl.height           = 10
$keywordFileLbl.location         = New-Object System.Drawing.Point(39,160)
$keywordFileLbl.Font             = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$keywordFileLbl.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$keywordTxtBox                   = New-Object system.Windows.Forms.TextBox
$keywordTxtBox.multiline         = $false
$keywordTxtBox.width             = 646
$keywordTxtBox.height            = 20
$keywordTxtBox.location          = New-Object System.Drawing.Point(22,188)
$keywordTxtBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$fileExtLbl                      = New-Object system.Windows.Forms.Label
$fileExtLbl.text                 = "Specific file extension"
$fileExtLbl.AutoSize             = $true
$fileExtLbl.width                = 25
$fileExtLbl.height               = 10
$fileExtLbl.location             = New-Object System.Drawing.Point(39,220)
$fileExtLbl.Font                 = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$fileExtLbl.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$fileExtLbl.Tool

$fileExt                   = New-Object system.Windows.Forms.TextBox
$fileExt.multiline         = $false
$fileExt.width             = 50
$fileExt.height            = 20
$fileExt.Text              = ".*"
$fileExt.location          = New-Object System.Drawing.Point(200,220)
$fileExt.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SaveConfigBtn                   = New-Object system.Windows.Forms.Button
$SaveConfigBtn.text              = "Save Config"
$SaveConfigBtn.width             = 100
$SaveConfigBtn.height            = 30
$SaveConfigBtn.location          = New-Object System.Drawing.Point(579,250)
$SaveConfigBtn.Font              = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$SaveConfigBtn.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffaaaa")


$ToolTip                         = New-Object system.Windows.Forms.ToolTip
$ToolTip.ToolTipTitle            = "Additional Help"
$ToolTip.isBalloon               = $false

$ToolTip.SetToolTip($keywordFileLbl,'[OPTIONAL] For example: Daily SQL, Standard (this will select all files that have these keywords in it)')
$ToolTip.SetToolTip($originalLbl, '[MANDATORY] Enter the Original folder path.')
$ToolTip.SetToolTip($inProgressLbl, '[MANDATORY] Enter the In Progress folder path.')
$ToolTip.SetToolTip($jobLbl, '[MANDATORY] This is the automation name that you want to associate to these files/folder. If you do not want the program to automatically move your files to the InProgress folder, add the statement "{ignore-path}" to the end of your function name (no spaces.)')
$ToolTip.SetToolTip($completeLbl, '[OPTIONAL] This is optional, if you want to allow your code to send files to the completed folder after a certain condition.')
$ToolTip.SetToolTip($AddModuleBtn, 'Once you are done with a job, click Add Module to save it. You can enter more modules after that.')
$ToolTip.SetToolTip($SaveConfigBtn, 'Click here to save your configuration file. You can close the program after it saves.')
$ToolTip.SetToolTip($fileExtLbl, 'If you want the automation to run on specific file types. If all files, leave as ".*" NOTE: Only .xlsx and .csv files can be processed through excel automation. .CSV files will be converted to .XLSX. ')

$Form.controls.AddRange(@($jobTxtBox,$jobLbl,$originalLbl,$originalTxtBox,$AddModuleBtn,
    $completeLbl,$completeTxtBox,$inProgressLbl,$inProgressTxtBox,$Button1,
    $SaveConfigBtn, $keywordFileLbl, $keywordTxtBox, $ImportBtn, $fileExtLbl, $fileExt))

$AddModuleBtn.Add_Click({ AddModule })
$SaveConfigBtn.Add_Click({ SaveConfig })
$ImportBtn.Add_Click( { ImportConfig } )

function ImportConfig {
    $path = "$(Get-Location)\savedConfig.xml"
    Test-Path -Path $path
    if ((Test-Path -Path $path) -ne $true) {
        $ImportBtn.Text = "No config found."
        return
    }

    $global:config = Import-Clixml -Path "$(Get-Location)\savedConfig.xml"
    $ImportBtn.Text = "Imported!"
}
function SaveConfig {
    foreach ($i in 1..10) {
        $SaveConfigBtn.Text = "Building$('.' * $i)"
        Start-Sleep -Milliseconds 50
    }
    $config | Export-Clixml -Path ("$(Get-Location)\savedConfig.xml")
    $SaveConfigBtn.Text = "Save Config" 
}

function AddModule {
    $config[$jobTxtBox.Text] = @{original=$originalTxtBox.Text 
        inProgress=$inProgressTxtBox.Text
        completed=$completeTxtBox.Text
        fileNames=$keywordTxtBox.Text
        fileType=($fileExt.Text -split ",").Trim()
    }
    $jobTxtBox.Clear()
    $originalTxtBox.Clear()
    $inProgressTxtBox.Clear()
    $completeTxtBox.Clear()
    $keywordTxtBox.Clear()
    $fileExt.Clear()
 }


#Write your logic code here

[void]$Form.ShowDialog()