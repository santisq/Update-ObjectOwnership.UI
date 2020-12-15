function setOwner{
param(
    [string]$Object,
    [string]$Owner
)

$domainVal=$Owner.Split('\\')[0]
$ownerVal=$Owner.Split('\\')[1]

$obj=Get-ADObject $Object|select -First 1
$objPath="AD:{0}" -f $obj.DistinguishedName
$ACL=Get-Acl -Path $objPath
$ownerObj=New-Object System.Security.Principal.NTAccount($domainVal,$ownerVal)
$ACL.SetOwner($ownerObj)
Set-Acl -Path $objPath -AclObject $ACL

}