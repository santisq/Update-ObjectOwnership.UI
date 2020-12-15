$message="This will update the Owner value of the objects listed on the DataGrid.

Do you want to proceed?"

$hit1=Show-MessageBox -Message $message -Icon Question -Buttons YesNo

if($hit1 -eq 'Yes'){

$NewOwner=($Grid.NewOwner)[0]
$setCount=($Grid|?{$_.Status -eq 'Ready to Process'}).count

$message="{0} will be set as Owner of {1} objects.

Are you sure you wish to proceed?" -f $NewOwner,$setCount

$hit2=Show-MessageBox -Message $message -Icon Warning -Buttons YesNo

}

if($hit2 -eq 'Yes'){

$dataGrid.Rows.Clear()
$progressBar.Value=0
$label.Text='Processing...'
$logOut=New-Object System.Collections.ArrayList

if($Grid.Status -eq 'Ready to Process'){

    $i=0;foreach($item in $Global:Grid){
    
        $i++
        $percentage=$i/$Grid.count*100
        $progressBar.Value=$percentage
        $form.Refresh()
        sleep -Milliseconds 1
    
        if($item.Status -eq 'Ready to Process'){
            try{
                
                setOwner -Object $item.ObjectGUID -Owner $item.NewOwner
                $item.Status='Updated'
                $logOut.Add(
                    [pscustomobject]@{
                        CanonicalName=$item.CanonicalName
                        ObjectGUID=$item.ObjectGUID
                        ProcessedBy=$env:USERNAME
                        OldOwner=$item.CurrentOwner
                        NewOwner=$item.NewOwner
                        Status=$item.Status
                    }
                ) > $null

            }catch{

                $item.Status=$_
                $logOut.Add(
                    [pscustomobject]@{
                        CanonicalName=$item.CanonicalName
                        ObjectGUID=$item.ObjectGUID
                        ProcessedBy=$env:USERNAME
                        OldOwner=$item.CurrentOwner
                        NewOwner=$item.NewOwner
                        Status=$item.Status
                    }
                ) > $null

            }
            
            $dataGrid.Rows.Add(
                $item.CanonicalName,
                $item.ObjectGUID,
                $item.CurrentOwner,
                $item.NewOwner,
                $item.Status
            )

        }else{

           $dataGrid.Rows.Add(
                $item.CanonicalName,
                $item.ObjectGUID,
                $item.CurrentOwner,
                $item.NewOwner,
                $item.Status
            )
        }
    }
}

if($logOut){

    $logPath=Split-Path $PSScriptRoot
    $date=([datetime]::Now).ToString('MM-dd-yyyy')
    $fileName="Set Owner Logs - {0}.csv" -f ([datetime]::Now).ToString('HHmmss')
    $logPath=Join-Path $logPath -ChildPath "Logs\$date"

    if(!(Test-Path $logPath)){New-Item -Path $logPath -ItemType Directory > $null}

    $logOut|Export-Csv -Path "$logPath\$fileName" -NoTypeInformation

}

$progressBar.Value=0
$label.Text='Ready'
$processButton.Enabled=$false
$processButton.Enabled=$false
$browseButton.Enabled=$false
$changeOwnerButton.Enabled=$false

}