Import-Csv -Path C:\KSMC\UserChangesTest.csv | Foreach-Object {
    $user = $_.UserPrincipalName
    $title = $_.Title
    $department = $_.Department
    Get-ADUser -Filter 'UserPrincipalName -eq $user' | Set-ADUser -Title $title -Department $department
}