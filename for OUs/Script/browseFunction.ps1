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

if($obj -and $obj[0].psObject.Properties.Name -match 'CanonicalName'){
    
    $OUs=$obj.CanonicalName
    $label.Text='Processing...'

    rv obj

    $i=0;foreach($OU in $OUs){
        
        if($OU -notin $Grid.CanonicalName){
            
            $i++
            $percentage=$i/$OUs.count*100
            $progressBar.Value=$percentage
            $form.Refresh()
            sleep -Milliseconds 1

            try{
                $ouName=$OU.Split('/')[-1]
                $obj=Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -Properties nTSecurityDescriptor,CanonicalName
                $obj=$obj|?{$_.CanonicalName -eq $OU}
            
                $Global:Grid.Add(
                    [pscustomobject]@{
                        CanonicalName=$obj.CanonicalName
                        ObjectGUID=$obj.ObjectGUID.ToString().ToUpper()
                        CurrentOwner=$obj.nTSecurityDescriptor.Owner
                        NewOwner=('{0}\{1}' -f $domain,$owner.samAccountName)
                        Status='Ready to Process'
                }) > $null

                $dataGrid.Rows.Add(
                    $Grid[-1].CanonicalName,
                    $Grid[-1].ObjectGUID,
                    $Grid[-1].CurrentOwner,
                    $Grid[-1].NewOwner,
                    $Grid[-1].Status
                )
            
            }catch{

                $Global:Grid.Add(
                    [pscustomobject]@{
                        CanonicalName=$OU
                        ObjectGUID='-'
                        CurrentOwner='-'
                        NewOwner='-'
                        Status='Not Found'
                    }) > $null   

                $dataGrid.Rows.Add(
                    $Grid[-1].CanonicalName,
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

}elseif($obj -and $obj[0].psObject.Properties.Name -notmatch 'CanonicalName'){
    $message="Excel/Csv file must contain a column with Name 'CanonicalName'."
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