function OKbtnEventListener($NewOwner){

$formChg.Dispose()
$domain=(Get-ADDomain).NetBiosName
$ownerLabelValue.Text=$newOwnerVal.DistinguishedName
$Global:defaultOwnerValue=$newOwnerVal
$dataGrid.Rows.Clear()
$label.Text='Processing...'

if($Grid){
    $i=0;foreach($item in $Global:Grid){
        $i++
        $percentage=$i/$Grid.count*100
        $progressBar.Value=$percentage
        $form.Refresh()
        sleep -Milliseconds 1
        if($item.Status -match 'Ready to Process|Updated'){
            $item.NewOwner=('{0}\{1}' -f $domain,$defaultOwnervalue.samAccountName)
            $dataGrid.Rows.Add(
                $item.'GPO Name',
                $item.'GPO ID',
                $item.CurrentOwner,
                $item.NewOwner,
                $item.Status
            )
        }else{
            $dataGrid.Rows.Add(
                $item.'GPO Name',
                $item.'GPO ID',
                $item.CurrentOwner,
                $item.NewOwner,
                $item.Status
            )
        }            
    }
}

$label.Text='Ready...'
$progressBar.Value=0

}

OKbtnEventListener -NewOwner $newOwnerVal