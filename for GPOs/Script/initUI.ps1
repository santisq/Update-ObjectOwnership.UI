function Use-RunAs{ 
param([Switch]$Check) 

#Thanks to Matt Painter for this function.

$IsAdmin=([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")

    if ($Check) { return $IsAdmin }     
 
    if ($MyInvocation.ScriptName -ne "") 
    {  
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $arg = "-file `"$($MyInvocation.ScriptName)`""
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -EA 'Stop' -WindowStyle Maximized
            } 
            catch 
            { 
                Write-Warning "Error - Failed to restart script with runas"  
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }  
    else  
    {  
        Write-Warning "Error - Script must be saved as a .ps1 file first"  
        break  
    }  
}
 
Use-RunAs

. "$PSScriptRoot\Show-Message.ps1"
. "$PSScriptRoot\setOwnerFunc.ps1"

try{

Import-Module ImportExcel, ActiveDirectory

$Global:Grid=New-Object System.Collections.ArrayList

$ErrorActionPreference='Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
[Windows.Forms.Application]::EnableVisualStyles()

$form=New-Object System.Windows.Forms.Form
#$form.ClientSize=New-Object System.Drawing.Size($resolution.Width,$resolution.Height)
$form.StartPosition='CenterScreen'
$form.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon("$PSHOME\PowerShell.exe")
#$form.FormBorderStyle='Fixed3D'
#$form.AutoScalemode='DPI'
$form.KeyPreview=$True
$form.Top=$true
$form.TopMost=$true
$form.Text= 'Set GPO Ownership'
$form.MaximizeBox=$false
$form.WindowState='Maximized'

$bounds=($form.CreateGraphics()).VisibleClipBounds|select Width,Height
$height=$bounds.Height
$width=$bounds.Width

$dataGrid=New-Object System.Windows.Forms.DataGridView
$dataGrid.Size=New-Object System.Drawing.Size($($width-20),$($height-150))
$dataGrid.Location=New-Object System.Drawing.Size(10,60)
$dataGrid.SelectionMode='FullRowSelect'
$dataGrid.MultiSelect=$false
$dataGrid.Font=New-Object System.Drawing.Font('Calibri',10,[System.Drawing.FontStyle]::Regular)
$dataGrid.RowTemplate.Height=20
$dataGrid.ColumnHeadersHeight=30
$dataGrid.ReadOnly=$true
$datagrid.Anchor = 'Top, Bottom, Left'
$form.Controls.Add($dataGrid)

$buttonHeightLocation=$dataGrid.Size.Height+70

$processButton=New-Object System.Windows.Forms.Button
$processButton.Size=New-Object System.Drawing.Size(85,35)
$processButton.Location=New-Object System.Drawing.Size($($width-95),$buttonHeightLocation)
$processButton.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$processButton.Text="&Process"
$processButton.Enabled=$false
$processButton.Add_Click({

. "$PSScriptRoot\processFunction.ps1" 

})
$form.Controls.Add($processButton)

$browseButton=New-Object System.Windows.Forms.Button
$browseButton.Size=New-Object System.Drawing.Size(85,35)
$browseButton.Location=New-Object System.Drawing.Size($($processButton.Location.X-$processButton.Size.Width-5),$buttonHeightLocation)
$browseButton.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$browseButton.Text="&Browse"

$progressBar=New-Object System.Windows.Forms.ProgressBar
$progressBar.Name='progressBar'
$progressBar.Value=0
$progressBar.Style='Continuous'
$progressBar.Size=New-Object System.Drawing.Size($($width-20),20)
$progressBar.Location=New-Object System.Drawing.Size(10,$($height-30))
$form.Controls.Add($progressBar)

$label=New-Object System.Windows.Forms.Label
$label.Name='status'
$label.Text='Ready'
$label.Font=New-Object System.Drawing.Font('Calibri',10,[System.Drawing.FontStyle]::Regular)
$label.BackColor=[System.Drawing.Color]::FromName('Transparent')
$label.Size=New-Object System.Drawing.Size(200,30)
$label.Location=New-Object System.Drawing.Size(10,$($height-50))
$form.Controls.Add($label)

$ownerLabel=New-Object System.Windows.Forms.Label
$ownerLabel.Name='owner'
$ownerLabel.Text='New Owner Value'
$ownerLabel.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$ownerLabel.BackColor=[System.Drawing.Color]::FromName('Transparent')
$ownerLabel.Size=New-Object System.Drawing.Size(200,15)
$ownerLabel.Location=New-Object System.Drawing.Size(10,5)
$form.Controls.Add($ownerLabel)

$Global:defaultOwnerValue=Get-ADObject -LDAPFilter "(samAccountName=Domain Admins)" -Properties samAccountName

$ownerLabelValue=New-Object System.Windows.Forms.Label
$ownerLabelValue.Name='ownervalue'
$ownerLabelValue.Text=$defaultOwnerValue.DistinguishedName
$ownerLabelValue.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Bold)
$ownerLabelValue.Padding='10,5,2,2'
$ownerLabelValue.BorderStyle=2
$ownerLabelValue.Size=New-Object System.Drawing.Size($($width-100),30)
$ownerLabelValue.Location=New-Object System.Drawing.Size(10,23)
$form.Controls.Add($ownerLabelValue)

$changeOwnerButton=New-Object System.Windows.Forms.Button
$changeOwnerButton.Size=New-Object System.Drawing.Size(75,35)
$changeOwnerButton.Location=New-Object System.Drawing.Size($($width-85),20)
$changeOwnerButton.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$changeOwnerButton.Text="&Change"
$changeOwnerButton.Add_Click({

. "$PSScriptRoot\changeBtn.ps1"

})
$form.Controls.Add($changeOwnerButton)



$colums=@(
'GPO Name','GPO ID'
'Current Owner','New Owner'
'Status'
)

$dataGrid.ColumnCount=$colums.Count
$dataGrid.ColumnHeadersVisible=$true

$i=0;$colums|%{
    $dataGrid.Columns[$i].Name=$_
    $i++
}
$dataGrid.AutoSizeColumnsMode=[System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$dataGrid.Columns[1].Width=280

$browseButton.Add_Click({

. "$PSScriptRoot\browseFunction.ps1"

})
$form.Controls.Add($browseButton)

$clearButton=New-Object System.Windows.Forms.Button
$clearButton.Size=New-Object System.Drawing.Size(85,35)
$clearButton.Location=New-Object System.Drawing.Size($($browseButton.Location.X-$browseButton.Size.Width-5),$buttonHeightLocation)
$clearButton.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$clearButton.Text="C&lear"
$clearButton.Enabled=$false
$clearButton.Add_Click({
    $Global:Grid=New-Object System.Collections.ArrayList
    $dataGrid.Rows.Clear()
    $clearButton.Enabled=$false
    $browseButton.Enabled=$True
    $changeOwnerButton.Enabled=$true
    $processButton.Enabled=$True
})
$form.Controls.Add($clearButton)

$helpbutton=New-Object System.Windows.Forms.Button
$helpbutton.Size=New-Object System.Drawing.Size(85,35)
$helpbutton.Location=New-Object System.Drawing.Size($($clearButton.Location.X-$clearButton.Size.Width-5),$buttonHeightLocation)
$helpbutton.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$helpbutton.Text='Help'
$helpbutton.Add_Click({

. "$PSScriptRoot\helpForm.ps1"

})
$form.Controls.Add($helpbutton)

$browseButton.Select()

$form.Add_KeyDown({
    
    if($_.KeyCode -eq 'Enter'){
        $browseButton.PerformClick()
    }
})

$form.Add_KeyDown({
    
    if($_.KeyCode -eq 'Escape'){
        $form.Dispose()
    }
})

$form.ShowDialog() > $null

}catch{

Show-MessageBox -Title 'Execution Error' -Buttons OK -Icon Error -Message $_

}