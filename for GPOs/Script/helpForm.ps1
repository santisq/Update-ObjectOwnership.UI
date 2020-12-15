$helpForm=New-Object System.Windows.Forms.Form
$helpForm.ClientSize=New-Object System.Drawing.Size(950,520)
$helpForm.StartPosition='CenterScreen'
$helpForm.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon("$PSHOME\PowerShell.exe")
$helpForm.FormBorderStyle='Fixed3D'
$helpForm.AutoScalemode='DPI'
$helpForm.KeyPreview=$True
$helpForm.TopMost=$true
$helpForm.Text= 'Help'
$helpForm.MaximizeBox=$false

$helpTxtBox=New-Object System.Windows.Forms.TextBox
$helpTxtBox.Text=$(cat "$psscriptroot\Help.txt"|out-string)
$helpTxtBox.Multiline=$True
$helpTxtBox.ReadOnly=$True
$helpTxtBox.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Bold)
$helpTxtBox.Size=New-Object System.Drawing.Size(930,450)
$helpTxtBox.Location=New-Object System.Drawing.Size(10,10)


$helpBtn=New-Object System.Windows.Forms.Button
$helpBtn.Size=New-Object System.Drawing.Size(75,30)
$helpBtn.Location=New-Object System.Drawing.Size($($helpForm.Width-100),$($helpForm.Height-80))
$helpBtn.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$helpBtn.Text="Close"
$helpBtn.Enabled=$True
$helpBtn.Add_Click({$helpForm.Dispose()})

$helpForm.Controls.Add($helpBtn)
$helpForm.Controls.Add($helpTxtBox)
$helpBtn.Select()

$helpForm.ShowDialog()