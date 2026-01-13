
foreach ($layout in @{
"Devices" = @("Device Type","DeviceType","os","Make","Model")
}){
    $layout = $null; $field = $null;
    $layout = get-huduassetlayout -name $layout.key; $layout = $layout.asset_layout ?? $layout;
    $fields = $layout.fields | where-object { $_.label -in $layout.value }
    foreach ($field in $fields){
        write-host " Processing listify for asset layout $($layout.name) field $($field.label)" -ForegroundColor Cyan
        Listify-TextField -assetlayout $layout.key -fieldlabel $field.label
    }
}