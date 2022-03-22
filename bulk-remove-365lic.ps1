## Bulk Remove licenses ##
## Select Csv file
$csv = Get-ChildItem -Path C:\temp\Office365License\Remove\users.csv -File | Out-GridView -PassThru

## Import Csv
$users = Import-Csv $csv.FullName

## Select Account SKU to be removed
$accountSKU  = Get-MsolAccountSku | Select-Object AccountSkuId | Out-GridView -PassThru

## Loop through each user in the Csv
foreach($user in $users){
Write-Host "Removing $($accountSKU.AccountSkuId) licence from $($user.UserPrincipalName)" -ForegroundColor Yellow

## Remove licence
Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $accountSKU.AccountSkuId
}