################## Functions ####################

if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Write-Host "Module exists"
} 
else {
    Write-Host "Module does not exist"
	PAUSE
	EXIT
	}


$Retrieve_Click={
    $Retrieve.Enabled = $false
    $Retrieve.Text = 'Getting Results'
    $array = @()
    $groups = Get-ADGroupMember -Identity $GroupName.Text
    foreach($item in $groups){
        
        $user = Get-ADUser $item.SamAccountName -Properties Name,SamAccountName,Title,Department,mail
        $array += $user    
    }
    $array = $array | select -Property Name,SamAccountName,Title,Department,mail
    $Global:resource = New-Object System.Collections.ArrayList(,$array)
    $DataGridView1.DataSource=$Global:resource
    $DataGridView1.Columns | Foreach-Object{
        $_.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        }
    $DataGridView1.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    $Retrieve.Enabled = $true
    $Retrieve.text                   = "Retrieve"
}

Function Get-SaveFile($initialDirectory)
{
$tmpFileName = $GroupName.Text+$(Get-Date -Format "MMddyyyyHHmm")+'.csv'
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
Out-Null

$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.initialDirectory = $initialDirectory
$SaveFileDialog.FileName = "$tmpFileName"
$SaveFileDialog.filter = "All files (*.*)| *.*"
$SaveFileDialog.ShowDialog() | Out-Null
$resource | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation
}
#################################################




Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '821,767'
$Form.text                       = "Group Members Attributes"
$Form.TopMost                    = $false

$Retrieve                        = New-Object system.Windows.Forms.Button
$Retrieve.text                   = "Retrieve"
$Retrieve.width                  = 113
$Retrieve.height                 = 30
$Retrieve.Anchor                 = 'top,right'
$Retrieve.location               = New-Object System.Drawing.Point(524,40)
$Retrieve.Font                   = 'Microsoft Sans Serif,10'

$GroupName                       = New-Object system.Windows.Forms.TextBox
$GroupName.multiline             = $false
$GroupName.width                 = 483
$GroupName.height                = 40
$GroupName.Anchor                = 'top,right,left'
$GroupName.location              = New-Object System.Drawing.Point(18,37)
$GroupName.Font                  = 'Microsoft Sans Serif,15'

$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
$DataGridView1.Name              = "DataGridView1"
$DataGridView1.width             = 769
$DataGridView1.height            = 620
$DataGridView1.Anchor            = 'top,right,bottom,left'
$DataGridView1.location          = New-Object System.Drawing.Point(19,83)

$Save                            = New-Object system.Windows.Forms.Button
$Save.text                       = "Save"
$Save.width                      = 60
$Save.height                     = 30
$Save.Anchor                     = 'bottom'
$Save.location                   = New-Object System.Drawing.Point(370,722)
$Save.Font                       = 'Microsoft Sans Serif,10'

$Title                           = New-Object system.Windows.Forms.Label
$Title.text                      = "Group Name"
$Title.AutoSize                  = $true
$Title.width                     = 25
$Title.height                    = 10
$Title.location                  = New-Object System.Drawing.Point(26,15)
$Title.Font                      = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Retrieve,$GroupName,$DataGridView1,$Save,$Title))

$Retrieve.Add_Click($Retrieve_Click)
$Save.Add_Click({Get-SaveFile})


[void]$Form.ShowDialog()
EXIT