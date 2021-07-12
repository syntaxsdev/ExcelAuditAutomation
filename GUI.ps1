<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    bak
.SYNOPSIS
    Form
#>
$global:config = @{}

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
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

$IgnorePathChk                   = New-Object system.Windows.Forms.CheckBox
$IgnorePathChk.text              = "Ignores Default Path?"
$IgnorePathChk.AutoSize          = $false
$IgnorePathChk.width             = 200
$IgnorePathChk.height            = 20
$IgnorePathChk.location          = New-Object System.Drawing.Point(280, 220)
$IgnorePathChk.Font              = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$IgnorePathChk.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$SaveToXlChk                   = New-Object system.Windows.Forms.CheckBox
$SaveToXlChk.text              = "Convert to Excel if .csv?"
$SaveToXlChk.AutoSize          = $false
$SaveToXlChk.width             = 200
$SaveToXlChk.height            = 20
$SaveToXlChk.Checked           = $true
$SaveToXlChk.location          = New-Object System.Drawing.Point(480, 220)
$SaveToXlChk.Font              = New-Object System.Drawing.Font('Microsoft YaHei UI',10)
$SaveToXlChk.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

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
$fileExt.Text              = "*.csv"
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
$ToolTip.SetToolTip($SaveToXlChk, 'Select this if you have .csv files you want to convert to Excel. This will create a .xlsx version in the {IN PROGRESS} folder, leaving the original alone. If you want to use the original file or csv, untick this.')
$ToolTip.SetToolTip($IgnorePathChk, 'Select this if you plan to use the SaveQuitAndMove() method in your automation')
$Form.controls.AddRange(@($jobTxtBox,$jobLbl,$originalLbl,$originalTxtBox,$AddModuleBtn,
    $completeLbl,$completeTxtBox,$inProgressLbl,$inProgressTxtBox,$Button1,
    $SaveConfigBtn, $keywordFileLbl, $keywordTxtBox, $ImportBtn, $fileExtLbl, $fileExt, $IgnorePathChk, $SaveToXlChk))

$AddModuleBtn.Add_Click({ AddModule })
$SaveConfigBtn.Add_Click({ SaveConfig })
$ImportBtn.Add_Click( { ImportConfig } )
$IgnorePathChk.Add_Click({ IgnorePath })

function IgnorePath() {

}
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
    $jobName = $jobTxtBox.Text
    if ($IgnorePathChk.Checked) {
        $jobName = "$jobName{ignore-path}"
    }
    $config[$jobName] = @{original=$originalTxtBox.Text
        inProgress=$inProgressTxtBox.Text
        completed=$completeTxtBox.Text
        fileNames=$keywordTxtBox.Text
        saveToXl=$SaveToXlChk.Checked
        fileType=($fileExt.Text -split ",").Trim()
    }
    $jobTxtBox.Clear()
    $originalTxtBox.Clear()
    $inProgressTxtBox.Clear()
    $completeTxtBox.Clear()
    $keywordTxtBox.Clear()
    $fileExt.Text = "*.csv"
 }


#Write your logic code here

[void]$Form.ShowDialog()