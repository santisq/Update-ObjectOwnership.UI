function browseFunction($Owner){

$DN=Get-ADDomain
$domain=$DN.NetBiosName
$progressBar.Value=0

$browse = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = Join-Path (Split-Path $PSScriptRoot) -ChildPath 'Input'
    Filter = 'Microsoft Excel Open XML Spreadsheet (*.xlsx)|*.xlsx|Comma-separated Values (*.csv)|*.csv'
    Multiselect=$true
}
$browse.ShowDialog() > $null

try{
    if($browse.FileNames){
        switch -Regex($browse.FileNames){
            '\.xlsx$'{
                $obj=$browse.FileNames|%{Import-Excel $_}
            }
            '\.csv$'{
                $obj=$browse.FileNames|%{Import-Csv $_}
            }
        }
        rv browse
    }
}
catch{Write-Warning $_}

if($obj -and $obj[0].psObject.Properties.Name -match 'ComputerName'){
    
    $computers=$obj.ComputerName
    $label.Text='Processing...'

    rv obj

    $i=0;foreach($computer in $computers){
        
        if($computer -notin $Grid.ComputerName){
            
            $i++
            $percentage=$i/$computers.count*100
            $progressBar.Value=$percentage
            $form.Refresh()
            sleep -Milliseconds 1

            try{
            
                $obj=Get-ADComputer $computer -Properties nTSecurityDescriptor
            
                $Global:Grid.Add(
                    [pscustomobject]@{
                        ComputerName=$obj.Name
                        ObjectGUID=$obj.ObjectGUID.ToString().ToUpper()
                        CurrentOwner=$obj.nTSecurityDescriptor.Owner
                        NewOwner=('{0}\{1}' -f $domain,$owner.samAccountName)
                        Status='Ready to Process'
                }) > $null

                $dataGrid.Rows.Add(
                    $Grid[-1].ComputerName,
                    $Grid[-1].ObjectGUID,
                    $Grid[-1].CurrentOwner,
                    $Grid[-1].NewOwner,
                    $Grid[-1].Status
                )
            
            }catch{

                $Global:Grid.Add(
                    [pscustomobject]@{
                        ComputerName=$computer
                        ObjectGUID='-'
                        CurrentOwner='-'
                        NewOwner='-'
                        Status='Not Found'
                    }) > $null   

                $dataGrid.Rows.Add(
                    $Grid[-1].ComputerName,
                    $Grid[-1].ObjectGUID,
                    $Grid[-1].CurrentOwner,
                    $Grid[-1].NewOwner,
                    $Grid[-1].Status
                )     
            }
        }
    }

    $dataGrid.Columns|%{
        $_.AutoSizeMode=[System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    }
    
    $progressBar.Value=0

}elseif($obj -and $obj[0].psObject.Properties.Name -notmatch 'ComputerName'){
    $message="Excel/Csv file must contain a column with Name 'ComputerName'."
    Show-MessageBox -Message $message -Icon Error -Buttons OK
}

$dataGrid.ClearSelection()
$progressBar.Value=0
$label.Text='Ready'

if($Grid){
    $clearButton.Enabled=$true
}
if($Grid.Status -eq 'Ready to Process'){
    $processButton.Enabled=$true
}

}

browseFunction -Owner $defaultOwnerValue