$formChg=New-Object System.Windows.Forms.Form
$formChg.ClientSize=New-Object System.Drawing.Size(570,110)
$formChg.StartPosition='CenterScreen'
$formChg.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon("$PSHOME\PowerShell.exe")
$formChg.FormBorderStyle='Fixed3D'
$formChg.AutoScalemode='DPI'
$formChg.KeyPreview=$True
$formChg.TopMost=$true
$formChg.Text= 'Change Owner Value'
$formChg.MaximizeBox=$false

$labelChg=New-Object System.Windows.Forms.Label
$labelChg.Name='newownerlbl'
$labelChg.Text='Input the sAMAccountName or DistinguishedName of the New Owner and click Validate.'
$labelChg.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$labelChg.BackColor=[System.Drawing.Color]::FromName('Transparent')
$labelChg.Size=New-Object System.Drawing.Size(550,30)
$labelChg.Location=New-Object System.Drawing.Size(10,10)
$formChg.Controls.Add($labelChg)

$textBox=New-Object System.Windows.Forms.TextBox
$textBox.Name='newownerval'
$textBox.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Bold)
$textBox.Size=New-Object System.Drawing.Size(550,40)
$textBox.Location=New-Object System.Drawing.Size(10,40)
$formChg.Controls.Add($textBox)

$okBtn=New-Object System.Windows.Forms.Button
$okBtn.Size=New-Object System.Drawing.Size(75,30)
$okBtn.Location=New-Object System.Drawing.Size(485,70)
$okBtn.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$okBtn.Text='OK'
$okBtn.Enabled=$false
$okBtn.Add_Click({

. "$PSScriptRoot\OKbtnEventListener.ps1"

})

$cancelBtn=New-Object System.Windows.Forms.Button
$cancelBtn.Size=New-Object System.Drawing.Size(75,30)
$cancelBtn.Location=New-Object System.Drawing.Size($($okBtn.Location.X-77),70)
$cancelBtn.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$cancelBtn.Text='Cancel'
$cancelBtn.Add_Click({

    $formChg.Dispose()

})

$validateBtn=New-Object System.Windows.Forms.Button
$validateBtn.Size=New-Object System.Drawing.Size(75,30)
$validateBtn.Location=New-Object System.Drawing.Size($($cancelBtn.Location.X-77),70)
$validateBtn.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$validateBtn.Text='Validate'
$validateBtn.Add_Click({
    
    $Name=$textBox.Text
    $filter="(|(distinguishedname=$Name)(samaccountname=$Name)(name=$Name))"
    $newOwnerVal=Get-ADobject -LDAPFilter $filter -Properties samAccountName|select -First 1
    
    if($newOwnerVal){
        
        $textBox.Text=$newOwnerVal.DistinguishedName
        New-Variable -Name newOwnerVal -Value $newOwnerVal -Scope Script -Force
        $okBtn.Enabled=$true
        $validateBtn.Enabled=$false
        $textBox.Enabled=$false
        $clearBtn.Enabled=$True

    }else{
        
        $textBox.Text="'{0}' could not be found on Active Directory. Try again." -f $name

    }

})

$clearBtn=New-Object System.Windows.Forms.Button
$clearBtn.Size=New-Object System.Drawing.Size(75,30)
$clearBtn.Location=New-Object System.Drawing.Size($($validateBtn.Location.X-77),70)
$clearBtn.Font=New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Regular)
$clearBtn.Text='Clear'
$clearBtn.Enabled=$false
$clearBtn.Add_Click({

    $textBox.Clear()
    $textBox.Enabled=$True
    $validateBtn.Enabled=$True
    $okBtn.Enabled=$false

})


$formChg.Controls.Add($validateBtn)
$formChg.Controls.Add($okBtn)
$formChg.Controls.Add($cancelBtn)
$formChg.Controls.Add($clearBtn)

$formChg.ShowDialog()
$textBox.Focus()