
foreach ($layout in @{
"Devices" = @("os","Make","Model")
"Accounts" = @("Account Type")
"ConfigItems"  = @("Configuration Type","Vendor")

}){
    $layout = $null; $field = $null;
    $layout = get-huduassetlayouts -name $layout.key | select-object -first 1; $layout = $layout.asset_layout ?? $layout;
    $fields = $layout.fields | where-object { $_.label -in $layout.value }
    foreach ($field in $fields){
        write-host " Processing listify for asset layout $($layout.name) field $($field.label)" -ForegroundColor Cyan
        $layout = get-huduassetlayouts -name $layout.key | select-object -first 1; $layout = $layout.asset_layout ?? $layout;
        Listify-TextField -assetlayout $layout -fieldlabel $field.label
    }
}
