function browseFunction($Owner){

$DN=Get-ADDomain
$domain=$DN.NetBiosName
$PolicyContainer=$DN.SystemsContainer
$PolicyContainer="{0},$PolicyContainer" -f 'CN=Policies'
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

$cond1=$obj -and $obj[0].psObject.Properties.Name -match 'Path'
$cond2=$obj.Path -match $PolicyContainer

if($cond1 -and $cond2){
    
    $GPpaths=$obj.Path
    $label.Text='Processing...'

    rv obj

    $i=0;foreach($path in $GPpaths){
            
        $i++
        $percentage=$i/$GPpaths.count*100
        $progressBar.Value=$percentage
        $form.Refresh()
        sleep -Milliseconds 1

        try{
            
            $gpo=Get-ADObject $path -Properties nTSecurityDescriptor,DisplayName
            
            if($gpo.Name -notin $Grid.'GPO ID'){
            
                $Global:Grid.Add(
                    [pscustomobject]@{
                        'GPO Name'=$gpo.DisplayName
                        'GPO ID'=$gpo.Name
                        CurrentOwner=$gpo.nTSecurityDescriptor.Owner
                        NewOwner=('{0}\{1}' -f $domain,$owner.samAccountName)
                        Status='Ready to Process'
                }) > $null

                $dataGrid.Rows.Add(
                    $Grid[-1].'GPO Name',
                    $Grid[-1].'GPO ID',
                    $Grid[-1].CurrentOwner,
                    $Grid[-1].NewOwner,
                    $Grid[-1].Status
                )
            }

        }catch{

            $Global:Grid.Add(
                [pscustomobject]@{
                    'GPO Name'='-'
                    'GPO ID'=$path.Split(',')[0] -replace 'CN='
                    CurrentOwner='-'
                    NewOwner='-'
                    Status='Not Found'
                }) > $null   

            $dataGrid.Rows.Add(
                $Grid[-1].'GPO Name',
                $Grid[-1].'GPO ID',
                $Grid[-1].CurrentOwner,
                $Grid[-1].NewOwner,
                $Grid[-1].Status
            )     
        }
    }

    $dataGrid.Columns|%{
        $_.AutoSizeMode=[System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    }
    
    $progressBar.Value=0

}elseif($gpo -and $gpo[0].psObject.Properties.Name -notmatch 'Path'){
    $message="Excel/Csv file must contain a column with Name 'Path'."
    Show-MessageBox -Message $message -Icon Error -Buttons OK
}else{
    $example=Get-ADobject -filter * -SearchBase $PolicyContainer -SearchScope OneLevel|select -First 1
    $message="Path column must contain the paths of the GPOs.
Example: '{0}'" -f $example
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